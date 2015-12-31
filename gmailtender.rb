#!/usr/bin/ruby

# Gmail API: https://developers.google.com/gmail/api/v1/reference/users/messages#resource

require 'google/api_client'
require 'google/api_client/client_secrets'
require 'google/api_client/auth/installed_app'
require 'google/api_client/auth/storage'
require 'google/api_client/auth/storages/file_store'
require 'fileutils'
require 'logger'
require "net/http"
require "uri"
require 'base64'

LOGFILE = "/home/jeff/.gmailtender.log"

APPLICATION_NAME = 'Gmail Tender'
CLIENT_SECRETS_PATH = 'client_secret.json'
CREDENTIALS_PATH = File.join(Dir.home, '.credentials', "gmailtender.json")
SCOPE = 'https://www.googleapis.com/auth/gmail.modify'

##
# Ensure valid credentials, either by restoring from the saved credentials
# files or intitiating an OAuth2 authorization request via InstalledAppFlow.
# If authorization is required, the user's default browser will be launched
# to approve the request.
#
# @return [Signet::OAuth2::Client] OAuth2 credentials
def authorize
  FileUtils.mkdir_p(File.dirname(CREDENTIALS_PATH))

  file_store = Google::APIClient::FileStore.new(CREDENTIALS_PATH)
  storage = Google::APIClient::Storage.new(file_store)
  auth = storage.authorize

  if auth.nil? || (auth.expired? && auth.refresh_token.nil?)
    app_info = Google::APIClient::ClientSecrets.load(CLIENT_SECRETS_PATH)
    flow = Google::APIClient::InstalledAppFlow.new({
      :port => 9293,
      :client_id => app_info.client_id,
      :client_secret => app_info.client_secret,
      :scope => SCOPE})
    auth = flow.authorize(storage)
    $logger.info "credentials saved to #{CREDENTIALS_PATH}" unless auth.nil?
  end
  auth
end


def encodeURIcomponent str
  return URI.escape(str, Regexp.new("[^#{URI::PATTERN::UNRESERVED}]"))
end


def make_org_entry heading, context, priority, date, body
  heading.downcase!
  $logger.debug "TODO [#{priority}] #{heading} #{context}"
  $logger.debug "SCHEDULED: #{date}"
  $logger.debug "#{body}"
  title = encodeURIcomponent "[#{priority}] #{heading}  :#{context}:"
  body  = encodeURIcomponent "SCHEDULED: #{date}\n#{body}"
  uri = URI.parse("http://carbon:3333")
  http = Net::HTTP.new(uri.host, uri.port)
  request = Net::HTTP::Get.new("/capture/b/LINK/#{title}/#{body}")
  response = http.request(request)
  return response
end


def archive message
  $logger.info "archiving #{message.id}"
  Client.execute!(
    :api_method => Gmail_api.users.messages.modify,
    :parameters => { 'userId' => 'me', 'id' => message.id },
    :body_object => { 'removeLabelIds' => ['INBOX'] })
end


def get_body message
  results = Client.execute!(
    :api_method => Gmail_api.users.messages.get,
    :parameters => { :userId => 'me', :id => message.id})
  message_json = JSON.parse(results.data.to_json())
  mime_data = Base64.urlsafe_decode64(message_json['payload']['body']['data'])
  return mime_data
end


def process_transfer message, headers
  $logger.info "(#{__method__})"
  content_raw = Client.execute!(
    :api_method => Gmail_api.users.messages.get,
    :parameters => { :userId => 'me', :id => message.id, :format => 'raw'})
  # e.g:
  #  Amount: $39.99
  #  From: Orange Parker Allowance, XXXXXX1099
  #  To: Orange Checking, XXXXXX6515
  #  Memo: game
  #  Transferred On: 08/22/2015
  detail = content_raw.data.raw[/(Amount:.*?Transferred On:.*?\n)/m, 1]
  detail.gsub!("\015", '')
  response = make_org_entry 'capital one transfer money notice', 'capitalone:@quicken', '#C',
                            "<#{Time.now.strftime('%F %a')}>",
                            detail + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  if (response.code == '200')
    archive message
  else
    $logger.error("make_org_entry gave response @{response.code} @{response.message}")
  end
end


def process_capitalone_statement message, headers
  $logger.info "(#{__method__})"
  detail = ''
  response = make_org_entry 'account statement available', 'capitalone:@quicken', '#C',
                            "<#{Time.now.strftime('%F %a')}>",
                            detail +
                            "https://secure.capitalone360.com/myaccount/banking/login.vm#" +
                            "https://mail.google.com/mail/u/0/#inbox/#{message.id}\n"
  if (response.code == '200')
    archive message
  else
    $logger.error("make_org_entry gave response @{response.code} @{response.message}")
  end
end


def process_pershing_statement message, headers
  $logger.info "(#{__method__})"
  detail = ''
  response = make_org_entry 'account statement available', 'pershing:@quicken', '#C',
                            "<#{Time.now.strftime('%F %a')}>",
                            detail + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  if (response.code == '200')
    archive message
  else
    $logger.error("make_org_entry gave response @{response.code} @{response.message}")
  end
end


def process_paypal_statement message, headers
  $logger.info "(#{__method__})"
  body = get_body message
  detail = '' + body[/<a.*?href="(.*?)".*?View Statement<\/a>/m, 1] + "\n"
  response = make_org_entry 'account statement available', 'paypal:@quicken', '#C',
                            "<#{Time.now.strftime('%F %a')}>",
                            detail + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  if (response.code == '200')
    archive message
  else
    $logger.error("make_org_entry gave response @{response.code} @{response.message}")
  end
end


def process_pge_statement message, headers
    $logger.info "(#{__method__})"
    detail = ''
    response = make_org_entry 'pg&e statement available', 'amex:@quicken', '#C',
                              "<#{Time.now.strftime('%F %a')}>",
                              detail + "http://www.pge.com/MyEnergy\nhttps://mail.google.com/mail/u/0/#inbox/#{message.id}"
    if (response.code == '200')
      archive message
    else
      $logger.error("make_org_entry gave response @{response.code} @{response.message}")
    end
end


def process_chase_mortgage_statement message, headers
  $logger.info "(#{__method__})"
  detail = 'https://stmts.chase.com/stmtslist?AI=475283320'
  response = make_org_entry 'chase mortgage statement available', 'chase:@quicken', '#C',
                            "<#{Time.now.strftime('%F %a')}>",
                            detail + "\n" + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  if (response.code == '200')
    archive message
  else
    $logger.error("make_org_entry gave response @{response.code} @{response.message}")
  end
end


def process_chase_credic_card_statement message, headers
  $logger.info "(#{__method__})"
  detail = 'https://stmts.chase.com/stmtslist?AI=16258879'
  response = make_org_entry 'chase credit card statement available', 'chase:@quicken', '#C',
                            "<#{Time.now.strftime('%F %a')}>",
                            detail + "\n" + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  if (response.code == '200')
    archive message
  else
    $logger.error("make_org_entry gave response @{response.code} @{response.message}")
  end
end


def process_peets_reload message, headers
  $logger.info "(#{__method__})"
  detail = '$50'
  response = make_org_entry 'peet\'s card reload order', 'amex:@quicken', '#C',
                            "<#{Time.now.strftime('%F %a')}>", detail
  if (response.code == '200')
    archive message
  else
    $logger.error("make_org_entry gave response @{response.code} @{response.message}")
  end
end


def process_comcast_bill message, headers
  $logger.info "(#{__method__})"
  detail = ''
  response = make_org_entry 'comcast bill ready', 'amex:@quicken', '#C',
                            "<#{Time.now.strftime('%F %a')}>",
                            detail + 'https://customer.xfinity.com/Secure/MyAccount/'
  if (response.code == '200')
    archive message
  else
    $logger.error("make_org_entry gave response @{response.code} @{response.message}")
  end
end


def process_american_express_statement message, headers
  $logger.info "(#{__method__})"
  detail = 'https://online.americanexpress.com/myca/statementimage/us/welcome.do?request_type=authreg_StatementCycles&Face=en_US&sorted_index=0'
  response = make_org_entry 'account statement available', 'amex:@quicken', '#C',
                            "<#{Time.now.strftime('%F %a')}>",
                            detail + "\n" + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  if (response.code == '200')
    archive message
  else
    $logger.error("make_org_entry gave response @{response.code} @{response.message}")
  end
end


def process_verizon_bill message, headers
  $logger.info "(#{__method__})"
  detail = 'https://ebillpay.verizonwireless.com/vzw/accountholder/mybill/BillingSummary.action'
  response = make_org_entry 'verizon bill available', 'amex:@quicken', '#C',
                            "<#{Time.now.strftime('%F %a')}>",
                            detail + "\n" + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  if (response.code == '200')
    archive message
  else
    $logger.error("make_org_entry gave response @{response.code} @{response.message}")
  end
end


def process_amazon_order message, headers
  $logger.info "(#{__method__})"
  results = Client.execute!(
    :api_method => Gmail_api.users.messages.get,
    :parameters => { :userId => 'me', :id => message.id})
  message_json = JSON.parse(results.data.to_json())
  mime_data = message_json['payload']['parts'][0]['parts'][0]['body']['data']
  body = Base64.urlsafe_decode64 mime_data
  order = headers['Subject'][/Your Amazon.com order of (.*)\./, 1]
  url = body[/View or manage your orders in Your Orders:\r\n?(https:.*?)\r\n/m, 1]
  delivery = body[/\s*Guaranteed delivery date:\r\n\s*(.*?)\r\n/m, 1]
  if delivery.nil?
    delivery = body[/\s*Estimated delivery date:\r\n\s*(.*?)-?\r\n/m, 1]
  end
  delivery = delivery.nil? ? Time.now : Date.parse(delivery)
  delivery = delivery.strftime('%F %a')
  total = body[/Order Total: (\$.*)\r\n/, 1]
  #$logger.info "#{order} #{url} #{delivery} #{total}"

  detail = "#{url}\n#{total}"
  response = make_org_entry "order of #{order}", 'amazon:@quicken', '#C',
                             "<#{Time.now.strftime('%F %a')}>",
                             detail + "\n" + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  if (response.code == '200')
    detail = url
    response = make_org_entry "delivery of #{order}", 'amazon:@waiting', '#C',
                              "<#{delivery}>",
                              detail + "\n" + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
    if (response.code == '200')
      archive message
      return
    end
  end
  $logger.error("make_org_entry gave response @{response.code} @{response.message}")
end


def process_amazon_video_order message, headers
  $logger.info "(#{__method__})"
  results = Client.execute!(
    :api_method => Gmail_api.users.messages.get,
    :parameters => { :userId => 'me', :id => message.id})
  message_json = JSON.parse(results.data.to_json())
  mime_data = message_json['payload']['parts'][0]['parts'][0]['body']['data']
  body = Base64.urlsafe_decode64 mime_data
  order = headers['Subject'][/Amazon.com order of (.*)\./, 1]
  total = body[/Grand Total:\s+(\$.*)\r\n/, 1]
  #$logger.info "#{order} #{total}"

  detail = "#{total}"
  response = make_org_entry "order of #{order}", 'amazon:@quicken', '#C',
                             "<#{Time.now.strftime('%F %a')}>",
                             detail + "\n" + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  if (response.code == '200')
    archive message
  else
    $logger.error("make_org_entry gave response @{response.code} @{response.message}")
  end
end


def process_amazon_subscribe_and_save message, headers
  $logger.info "(#{__method__})"
  detail = 'https://www.amazon.com/manageyoursubscription'
  response = make_org_entry 'review amazon subscribe and save delivery', 'amazon:@home', '#C',
                            "<#{Time.now.strftime('%F %a')}>",
                            detail + "\n" + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  if (response.code == '200')
    archive message
  else
    $logger.error("make_org_entry gave response @{response.code} @{response.message}")
  end
end


def process_etrade_statement message, headers
  $logger.info "(#{__method__})"
  detail = 'https://edoc.etrade.com/e/t/onlinedocs/docsearch?doc_type=stmt'
  response = make_org_entry 'etrade statement available', 'etrade:@quicken', '#C',
                            "<#{Time.now.strftime('%F %a')}>",
                            detail + "\n" + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  if (response.code == '200')
    archive message
  else
    $logger.error("make_org_entry gave response @{response.code} @{response.message}")
  end
end


def dispatch_message message, headers
  $logger.info headers['Subject']
  $logger.info headers['From']
  if headers['Subject'] == "Transfer Money Notice" &&
     headers['From'].include?("capitalone360.com")
    process_transfer message, headers
  elsif headers['Subject'].include?("eStatement's now available") &&
        headers['From'].include?("capitalone360.com")
    process_capitalone_statement message, headers
  elsif headers['Subject'] =='Brokerage Account Statement Notification' &&
        headers['From'] == '<pershing@advisor.netxinvestor.com>'
    process_pershing_statement message, headers
  elsif headers['Subject'].include?("account statement is available") &&
        headers['From'] == 'PayPal Statements <paypal@e.paypal.com>'
    process_paypal_statement message, headers
  elsif headers['Subject'] =='Your PG&E Energy Statement is Ready to View' &&
        headers['From'] == 'CustomerServiceOnline@pge.com'
    process_pge_statement message, headers
  elsif headers['Subject'] == 'Your mortgage statement is available online.' &&
        headers['From'] == 'Chase <no-reply@alertsp.chase.com>'
    process_chase_mortgage_statement message, headers
  elsif headers['Subject'] == 'Your credit card statement is available online' &&
        headers['From'] == 'Chase <no-reply@alertsp.chase.com>'
    process_chase_credic_card_statement message, headers
  elsif headers['Subject'].include?("Your Peet's Card Reload Order") &&
        headers['From'] == 'Customer Service <customerservice@peets.com>'
    process_peets_reload message, headers
  elsif headers['Subject'] == 'Your bill is ready' &&
        headers['From'] == 'XFINITY My Account <online.communications@alerts.comcast.net>'
    process_comcast_bill message, headers
  elsif headers['Subject'].index(/Important Notice: Your .* Statement/) &&
        headers['From'] == 'American Express <AmericanExpress@welcome.aexp.com>'
    process_american_express_statement message, headers
  elsif headers['Subject'] == 'Your Bill is Now Available' &&
        headers['From'] == 'Verizon Wireless <VZWMail@ecrmemail.verizonwireless.com>'
    process_verizon_bill message, headers
  elsif headers['Subject'].include?('Your Amazon.com order of') &&
        headers['From'] == '"auto-confirm@amazon.com" <auto-confirm@amazon.com>'
    process_amazon_order message, headers
  elsif headers['Subject'].include?('Amazon.com order of') &&
        headers['From'] == '"Amazon.com" <digital-no-reply@amazon.com>'
    process_amazon_video_order message, headers
  elsif headers['Subject'] == 'Amazon Subscribe & Save: Review Your Monthly Delivery' &&
        headers['From'] == '"Amazon Subscribe & Save" <no-reply@amazon.com>'
    process_amazon_subscribe_and_save message, headers
  elsif headers['Subject'] == 'You have a new account statement from E*TRADE Securities' &&
        headers['From'] == '"E*TRADE SECURITIES LLC" <etrade_stmt_mbox@statement.etradefinancial.com>'
    process_etrade_statement message, headers
  end
end


def redirect_output
  unless LOGFILE == 'STDOUT'
    logfile = File.expand_path(LOGFILE)
    FileUtils.mkdir_p(File.dirname(logfile), :mode => 0755)
    FileUtils.touch logfile
    File.chmod 0644, logfile
    $stdout.reopen logfile, 'a'
  end
  $stderr.reopen $stdout
  $stdout.sync = $stderr.sync = true
end


# setup logger
redirect_output unless $DEBUG

$logger = Logger.new STDOUT
$logger.level = $DEBUG ? Logger::DEBUG : Logger::INFO
$logger.info 'starting'

# Initialize the API
Client = Google::APIClient.new(:application_name => APPLICATION_NAME)
Client.authorization = authorize
Gmail_api = Client.discovered_api('gmail', 'v1')

# get the user's messages
results = Client.execute!(
  :api_method => Gmail_api.users.messages.list,
  :parameters => { :userId => 'me', :q => "in:inbox is:unread" })

$logger.info "#{results.data.messages.length} messages found"

results.data.messages.each { |message|
  # See https://developers.google.com/gmail/api/v1/reference/users/messages/get
  content = Client.execute!(
    :api_method => Gmail_api.users.messages.get,
    :parameters => { :userId => 'me', :id => message.id})
  $logger.debug "- #{message.id}"

  # See https://developers.google.com/gmail/api/v1/reference/users/messages#methods
  headers = {}
  content.data.payload.headers.each { |header|
    $logger.debug header.name
    headers[header.name] = header.value
  }

  dispatch_message message, headers
}

$logger.info 'done'
