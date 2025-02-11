#!/usr/bin/env ruby
# coding: utf-8     # rubocop:disable Style/Encoding
# frozen_string_literal: true

# GMail API: https://developers.google.com/gmail/api/quickstart/ruby
#            https://developers.google.com/gmail/api/v1/reference/
#            https://console.developers.google.com
# Google API Ruby Client:  https://github.com/google/google-api-ruby-client

require 'googleauth'
require 'googleauth/stores/file_token_store'
require 'google/apis/gmail_v1'
require 'google/apis/calendar_v3'

require 'chronic'
require 'fileutils'
require 'logger'
require 'net/http'
require 'nokogiri'
require 'addressable/uri'
require 'base64'
require 'thor'
require 'resolv-replace'

LOGFILE = File.join(Dir.home, '.log', 'gmailtender.log')

BASE_URI = 'https://www.google.com'
APPLICATION_NAME = 'gmailtender'
CLIENT_SECRETS_PATH = 'client_secret.json'
CREDENTIALS_PATH = File.join(Dir.home, '.credentials', 'gmailtender.yaml')
SCOPE = [Google::Apis::GmailV1::AUTH_GMAIL_MODIFY, 'https://www.googleapis.com/auth/calendar'].freeze

##
# Ensure valid credentials, either by restoring from the saved credentials
# files or intitiating an OAuth2 authorization. If authorization is required,
# the user's default browser will be launched to approve the request.
#
# @return [Google::Auth::UserRefreshCredentials] OAuth2 credentials
def authorize(interactive)
  FileUtils.mkdir_p(File.dirname(CREDENTIALS_PATH))
  client_id = Google::Auth::ClientId.from_file(CLIENT_SECRETS_PATH)
  token_store = Google::Auth::Stores::FileTokenStore.new(file: CREDENTIALS_PATH)
  authorizer = Google::Auth::UserAuthorizer.new(client_id, SCOPE, token_store, '')
  user_id = 'default'
  credentials = authorizer.get_credentials(user_id)
  if credentials.nil? && interactive
    url = authorizer.get_authorization_url(base_url: BASE_URI, scope: SCOPE)
    code = ask("Open the following URL in the browser and enter the resulting code after authorization\n#{url}")
    credentials = authorizer.get_and_store_credentials_from_code(
      user_id: user_id, code: code, base_url: BASE_URI
    )
  end
  credentials
end

class MessageHandler
  attr_accessor :gmail, :gcal, :options

  def initialize(params = {})
    params.each { |key, value| send "#{key}=", value }
  end

  def self.descendants
    ObjectSpace.each_object(Class).select { |klass| klass < self }
  end

  def make_org_entry(heading, context, priority, date, body)
    heading = heading.downcase
    $logger.debug "TODO [#{priority}] #{heading} #{context}"
    $logger.debug "SCHEDULED: #{date}"
    $logger.debug body.to_s
    title = Addressable::URI.encode_component "[#{priority}] #{heading}  :#{context}:", Addressable::URI::CharacterClasses::UNRESERVED
    body  = Addressable::URI.encode_component "SCHEDULED: #{date}\n#{body}", Addressable::URI::CharacterClasses::UNRESERVED
    uri = "http://cube.zt:3333/capture/b/LINK/#{title}/#{body}"
    $logger.info uri
    uri = Addressable::URI.parse uri

    http = Net::HTTP.new uri.host, uri.port
    request = Net::HTTP::Get.new uri.path

    return nil if options[:dry_run]

    response = http.request(request)
    $logger.error "make_org_entry gave response #{response.code} #{response.message} #{response.body}" if response.code != '200'

    response.code == '200'
  end

  def archive(message, label_name = 'INBOX')
    $logger.info "archiving #{message.id}"

    labels = (gmail.list_user_labels 'me').labels
    label_id = labels.select { |label| label.name == label_name }.first.id

    mmr = Google::Apis::GmailV1::ModifyMessageRequest.new(remove_label_ids: [label_id])
    gmail.modify_message 'me', message.id, mmr
  end

  def unlabel_thread(thread, label_name)
    $logger.info "archiving #{thread.id}"

    labels = (gmail.list_user_labels 'me').labels
    label_id = labels.select { |label| label.name == label_name }.first.id

    mmr = Google::Apis::GmailV1::ModifyThreadRequest.new(remove_label_ids: [label_id])
    gmail.modify_thread 'me', thread.id, mmr
  end

  def self.dispatch(gcal, gmail, options, message, headers)
    $logger.info headers['Subject']
    $logger.info headers['From']
    MessageHandler.descendants.each do |handler|
      $logger.debug "matching #{handler}"
      begin
        if handler.match headers
          handler.new(gcal: gcal, gmail: gmail, options: options).process message, headers
          break
        end
      rescue StandardError => e
        $logger.error e.message
        $logger.error e.backtrace.inspect
      end
    end
  end

  def process(message, headers)
    $logger.info "(#{self.class})"

    response = handle message, headers
    archive message if response
  end

  def self.friendly_name(addr)
    name = addr[/"?(.*?)"?\s</, 1] || addr[/<?(.*?)@/, 1]
    unless name.nil?
      name.downcase!
      name.gsub!(/[. ]/, '_')
    end

    name
  end

  def self.refile(gmail, options, context, thread, message, headers)
    $logger.info headers['Subject']
    $logger.info headers['From']

    full_context = context
    if context == '@waiting'
      to = friendly_name headers['To']
      to = friendly_name headers['From'] if to == 'jeff_kowalski'
      full_context = [to, context].join(':')
    end

    return if options[:dry_run]

    handler = MessageHandler.new(gmail: gmail, options: options)
    response = handler.make_org_entry headers['Subject'], full_context, '#C',
                                      "<#{Time.now.strftime('%F %a')}>",
                                      "https://mail.google.com/mail/u/0/#inbox/#{message.id}"

    handler.unlabel_thread thread, context if response
  end
end

class MH_AllstateBill < MessageHandler
  def self.match(headers)
    headers['Subject'] == 'Allstate: Your billing document is ready to view online' &&
      headers['From'] == 'Allstate My Account <allstate@trns01.allstate-email.com>'
  end

  def handle(message, _headers)
    raw = (gmail.get_user_message 'me', message.id, format: 'raw').raw
    # Policy Number: XXXX63191
    # Policy Type: Auto - private passenger voluntary
    # Document Title: Auto policy schedule for the Recurring Credit Card Pay Plan
    # Due Date: Payments scheduled for the 16th
    # Minimum Amount Due: See Schedule
    detail = raw[/(Policy Number:.*?Minimum Amount Due:.*?\n)/m, 1]
    detail.gsub!("\015", '')
    puts detail

    make_org_entry 'allstate bill available', '@quicken', '#C',
                   "<#{Time.now.strftime('%F %a')}>",
                   detail +
                   "https://myaccount.allstate.com/anon/login/login.aspx\n" \
                   "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end

class MH_LGTubClean < MessageHandler
  def self.match(headers)
    headers['Subject'] == 'It’s time to perform a tub clean on your LG Washing Machine' &&
      headers['From'] == 'lgcarecenter@lge.com'
  end

  def handle(message, _headers)
    make_org_entry 'perform tub clean', '@home', '#C',
                   "<#{Time.now.strftime('%F %a')}>",
                   "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end

class MH_SchwabStatement < MessageHandler
  def self.match(headers)
    headers['Subject']&.include?('Your account eStatement is available') &&
      headers['From'] == '"Charles Schwab & Co., Inc." <donotreply@mail.schwab.com>'
  end

  def handle(message, _headers)
    make_org_entry 'schwab statement available', 'schwab:@quicken', '#C',
                   "<#{Time.now.strftime('%F %a')}>",
                   "https://schwab.com/sa_reports\n" \
                   "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end

class MH_ChaseCheckingStatement < MessageHandler
  def self.match(headers)
    headers['Subject']&.include?('Your statement is ready for account ending in') &&
      headers['From'] == 'Chase <no-reply@alertsp.chase.com>'
  end

  def handle(message, _headers)
    make_org_entry 'chase checking statement available', 'chase:@quicken', '#C',
                   "<#{Time.now.strftime('%F %a')}>",
                   "https://secure.chase.com/web/auth/nav?navKey=requestStatementsAndDocuments&documentType=STATEMENTS&mode=documents\n" \
                   "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end

class MH_ChaseCreditCardStatement < MessageHandler
  def self.match(headers)
    headers['Subject'] == 'Your credit card statement is ready' &&
      headers['From'] == 'Chase <no-reply@alertsp.chase.com>'
  end

  def handle(message, _headers)
    make_org_entry 'chase credit card statement available', 'chase:@quicken', '#C',
                   "<#{Time.now.strftime('%F %a')}>",
                   "https://secure05a.chase.com/web/auth/dashboard#/dashboard/documents/myDocs/index;mode=accounts;documentType=STATEMENTS\n" \
                   "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end

class MH_CapitalOneTransfer < MessageHandler
  def self.match(headers)
    headers['Subject'] == 'Transfer Money Notice' &&
      headers['From'].include?('capitalone.com')
  end

  def handle(message, _headers)
    payload = (gmail.get_user_message 'me', message.id).payload
    body = payload.parts[0].body.data
    # e.g:
    #  Amount: $39.99
    #  From: Orange Parker Allowance, XXXXXX1099
    #  To: Orange Checking, XXXXXX6515
    #  Memo: game
    #  Transferred On: 08/22/2015
    detail = body[/(Amount:.*?Transferred On:.*?\n)/m, 1]
    detail.gsub!("\015", '')
    make_org_entry 'capital one transfer money notice', 'capitalone:@quicken', '#C',
                   "<#{Time.now.strftime('%F %a')}>",
                   detail + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end

class MH_EbmudBill < MessageHandler
  def self.match(headers)
    headers['Subject'] == 'Your EBMUD bill is available online.' &&
      headers['From'] == 'noreply@ebmud.com'
  end

  def handle(message, _headers)
    payload = (gmail.get_user_message 'me', message.id).payload
    body = payload.body.data
    detail = "#{body[%r{(Due\s+date:\s*[\d/]+)}, 1]}\n#{body[/(balance:\s*\$[\d.]+)/m, 1]}\n"
    make_org_entry 'ebmud bill available', 'orange_checking:@quicken', '#C',
                   "<#{Time.now.strftime('%F %a')}>",
                   detail +
                   "https://www.ebmud.com/customers/account/manage-your-account\n" \
                   "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end

class MH_CapitalOneStatement < MessageHandler
  def self.match(headers)
    headers['Subject']&.include?('Your latest bank statement is ready') &&
      headers['From'] == 'Capital One <capitalone@notification.capitalone.com>'
  end

  def handle(message, _headers)
    make_org_entry 'account statement available', 'capitalone:@quicken', '#C',
                   "<#{Time.now.strftime('%F %a')}>",
                   "https://verified.capitalone.com/auth/signin#/esignin?Product=360Bank\n" \
                   "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end

class MH_PaypalStatement < MessageHandler
  def self.match(headers)
    headers['Subject']&.include?('Your statement is now available') &&
      headers['From'] == 'PayPal Credit <customercare@paypal.com>'
  end

  def handle(message, _headers)
    payload = (gmail.get_user_message 'me', message.id).payload
    body = payload.body.data
    detail = "#{body[%r{<a.*?href="(.*?)".*?available online</a>}m, 1]}\n"
    make_org_entry 'account statement available', 'paypal:@quicken', '#C',
                   "<#{Time.now.strftime('%F %a')}>",
                   detail + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end

class MH_CloverReceipt < MessageHandler
  def self.match(headers)
    headers['Subject']&.include?('Your receipt from') &&
      headers['From']&.include?('<app@clover.com>')
  end

  def handle(message, headers)
    payload = (gmail.get_user_message 'me', message.id).payload
    body = payload.parts[0].body.data

    # $6.95 total

    amount = body[/(\$\d+\.\d+) total/, 1].to_s
    payee  = headers['Subject'][/Your receipt from (.*)/, 1].to_s.downcase
    make_org_entry "receipt for #{payee}", '@quicken', '#C',
                   "<#{Time.now.strftime('%F %a')}>",
                   "#{amount}\nhttps://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end

class MH_SquareReceipt < MessageHandler
  def self.match(headers)
    headers['Subject']&.include?('Receipt from') &&
      headers['From']&.include?('<receipts@messaging.squareup.com>')
  end

  def handle(message, _headers)
    raw = (gmail.get_user_message 'me', message.id, format: 'raw').raw
    body = raw

    # You paid $10.00 with your AMEX ending in 2008 to Pizza Politana on Aug 22

    amount  = body[/You paid (\$\d+\.\d+) with/, 1].to_s
    account = body[/with your (\S+) ending in/, 1].to_s.downcase
    payee   = body[/to (.*?) on/, 1].to_s.downcase
    make_org_entry "receipt for #{payee}", "#{account}:@quicken", '#C',
                   "<#{Time.now.strftime('%F %a')}>",
                   "#{amount}\nhttps://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end

class MH_PershingStatement < MessageHandler
  def self.match(headers)
    headers['Subject'] == 'Account Statement Notification' &&
      headers['From'] == 'edelivery@investor.pershing.com'
  end

  def handle(message, _headers)
    make_org_entry 'account statement available', 'pershing:@quicken', '#C',
                   "<#{Time.now.strftime('%F %a')}>",
                   "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end

class MH_PGEStatement < MessageHandler
  def self.match(headers)
    headers['Subject'] == 'Your PG&E Energy Statement is Ready to View' &&
      headers['From'] == 'CustomerServiceOnline@billpay.pge.com'
  end

  def handle(message, _headers)
    make_org_entry 'pg&e statement available', 'orange_checking:@quicken', '#C',
                   "<#{Time.now.strftime('%F %a')}>",
                   "http://www.pge.com/MyEnergy\nhttps://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end

class MH_ATT_Wireless_Bill < MessageHandler
  def self.match(headers)
    headers['Subject'] == 'Your AT&T wireless bill is ready to view'
  end

  def handle(message, _headers)
    make_org_entry 'at&t wireless bill ready', 'amex:@quicken', '#C',
                   "<#{Time.now.strftime('%F %a')}>",
                   "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end

class MH_SonicBill < MessageHandler
  def self.match(headers)
    headers['Subject'] == 'Payment Scheduled' &&
      headers['From'].include?('Sonic Billing')
  end

  def handle(message, _headers)
    make_org_entry 'sonic bill', 'visa:@quicken', '#C',
                   "<#{Time.now.strftime('%F %a')}>",
                   "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end

class MH_PeetsReload < MessageHandler
  def self.match(headers)
    headers['Subject']&.include?("Your Peet's Card Reload Order") &&
      headers['From'].include?('<customerservice@peets.com>')
  end

  def handle(message, _headers)
    make_org_entry 'peet\'s card reload order', ':@quicken', '#C',
                   "<#{Time.now.strftime('%F %a')}>",
                   "$50\n" \
                   "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end

class MH_AmericanExpressStatement < MessageHandler
  def self.match(headers)
    headers['Subject']&.index(/Your .* Statement/) &&
      headers['From'] == 'American Express <AmericanExpress@welcome.americanexpress.com>'
  end

  def handle(message, _headers)
    make_org_entry 'account statement available', 'amex:@quicken', '#C',
                   "<#{Time.now.strftime('%F %a')}>",
                   "https://www.americanexpress.com/en-us/account/login?DestPage=https%3A%2F%2Fglobal.americanexpress.com%2Factivity%2Fstatements\n" \
                   "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end

class MH_BNYMellonStatement < MessageHandler
  def self.match(headers)
    headers['Subject'] == 'BNY Mellon, N.A. - E-Statement Notification' &&
      headers['From'] == 'noreply@yourmortgageonline.com'
  end

  def handle(message, _headers)
    make_org_entry 'mortgage statement available', 'bny_mellon:@quicken', '#C',
                   "<#{Time.now.strftime('%F %a')}>",
                   "https://www.yourmortgageonline.com/documents/statements\n" \
                   "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end

class MH_GoogleFiStatement < MessageHandler
  def self.match(headers)
    headers['Subject'] == 'Your Google Fi monthly statement' &&
      headers['From'] == 'Google Payments <payments-noreply@google.com>'
  end

  def handle(message, _headers)
    payload = (gmail.get_user_message 'me', message.id).payload
    body = payload.parts[0].parts[0].body.data
    total = body.scan(/Your total is (\$\d+\.\d+)/)&.first&.first # Your total is $40.51
    date = body.scan(/Auto-payment is scheduled for (\w+ \d+, \d+)/)&.first&.first # Auto-payment is scheduled for December 21, 2020
    make_org_entry 'google fi statement', 'visa:@quicken', '#C',
                   "<#{Time.now.strftime('%F %a')}>",
                   "https://mail.google.com/mail/u/0/#inbox/#{message.id}\n#{total}\n#{date}"
  end
end

class MH_AmazonSubscribeAndSave < MessageHandler
  def self.match(headers)
    headers['Subject'] == 'Amazon Subscribe & Save: Review Your Monthly Delivery' &&
      headers['From'] == '"Amazon Subscribe & Save" <no-reply@amazon.com>'
  end

  def handle(message, _headers)
    make_org_entry 'review amazon subscribe and save delivery', 'amazon:@home', '#C',
                   "<#{Time.now.strftime('%F %a')}>",
                   "https://www.amazon.com/manageyoursubscription\n" \
                   "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end

class MH_AmazonOrder < MessageHandler
  def self.match(headers)
    (headers['Subject']&.index(/Your Amazon.*? order/) ||
     headers['Subject']&.index(/^Ordered:/)) &&
      headers['From'] == '"Amazon.com" <auto-confirm@amazon.com>'
  end

  def handle(message, headers)
    payload = (gmail.get_user_message 'me', message.id).payload
    body = payload.parts[0].body.data
    order = body[/You ordered\s+(".*?")\s*\.\r\n/m, 1] ||
            headers['Subject'][/Your Amazon.*? order of (.*)\./, 1] ||
            headers['Subject'][/Your Amazon.*? order (#.*)/, 1] ||
            headers['Subject'][/Ordered:\s+"(.*)"/, 1]
    url = body[/(https:.*order-details.*?)\r\n/m, 1]
    delivery = body[/\s*((:?(:?Guaranteed|Estimated) delivery date:\s*\r\n)|(:?Arriving))\s*(?<date>.*?)\r\n/m, :date]
    delivery = (delivery.nil? ? Time.now : Chronic.parse(delivery)).strftime('%F %a')
    total = body[/Total[:]?[\s\r\n]+\$?(?<total>.*?)[\s\r\n]/m, :total]
    $logger.info "#{order} #{url} #{delivery} #{total}"

    detail = "#{url}\n#{total}"
    response = make_org_entry "order of #{order}", 'amazon:@quicken', '#C',
                              "<#{Time.now.strftime('%F %a')}>",
                              "#{detail}\nhttps://mail.google.com/mail/u/0/#inbox/#{message.id}"
    if response
      detail = url
      response = make_org_entry "delivery of #{order}", 'amazon:@waiting', '#C',
                                "<#{delivery}>",
                                "#{detail}\nhttps://mail.google.com/mail/u/0/#inbox/#{message.id}"
    end

    response
  end
end

class MH_UPSMyChoice < MessageHandler
  def self.match(headers)
    (headers['Subject'].include?('UPS Update: Package Scheduled for Delivery') ||
     headers['Subject'].include?('UPS Ship Notification')) &&
      headers['From'] == 'UPS <mcinfo@ups.com>'
  end

  def handle(message, _headers)
    payload = (gmail.get_user_message 'me', message.id).payload
    body = payload.parts[0].body.data
    doc = Nokogiri::HTML(body)
    tracking = doc.at_css('#trackingNumber').content.strip
    delivery = doc.at_css('#deliveryDateTime')
    date_raw = delivery.children.first.text.strip
    date = date_raw.match(%r{\d+/\d+/\d+}).to_s
    times_raw = delivery.children.last.text.strip
    time = nil
    time = times_raw[3..] if times_raw.start_with?('by ')
    times = times_raw.split(' - ')

    if time
      expected = Time.strptime("#{date} #{time}", '%m/%d/%Y %I:%M %p')
      make_org_entry "ups delivery of #{tracking}", 'ups:@waiting', '#C',
                     "<#{expected.strftime('%F %a %H:%M')}>",
                     "https://www.ups.com/track?loc=null&tracknum=#{tracking}&requester=WT/trackdetails\n" \
                     "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
    elsif times
      expected = times.map { |t| Time.strptime("#{date} #{t}", '%m/%d/%Y %I:%M %p') }
      make_org_entry "ups delivery of #{tracking}", 'ups:@waiting', '#C',
                     "<#{expected[0].strftime('%F %a %H:%M')}>--<#{expected[1].strftime('%F %a %H:%M')}>",
                     "https://www.ups.com/track?loc=null&tracknum=#{tracking}&requester=WT/trackdetails\n" \
                     "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
    else
      expected = Date.strptime(date, '%m/%d/%Y')
      make_org_entry "ups delivery of #{tracking}", 'ups:@waiting', '#C',
                     "<#{expected.strftime('%F %a')}>",
                     "https://www.ups.com/track?loc=null&tracknum=#{tracking}&requester=WT/trackdetails\n" \
                     "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
    end
  end
end

class MH_USPSDelivery < MessageHandler
  def self.match(headers)
    headers['Subject']&.index(/USPS® (Expected|Scheduled) Delivery/) &&
      headers['From'] == 'auto-reply@usps.com'
  end

  def handle(message, headers)
    date, time, tracking = headers['Subject'].scan(/Delivery .. (.*) arriving by (.*?) ([A-Z0-9]+)/).first
    if time
      expected = Time.parse("#{date} #{time}")
      make_org_entry "usps delivery of #{tracking}", 'usps:@waiting', '#C',
                     "<#{expected.strftime('%F %a %H:%M')}>",
                     "https://tools.usps.com/go/TrackConfirmAction?tLabels=#{tracking}\n" \
                     "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
    else
      date, time1, time2, tracking = headers['Subject'].scan(/Delivery on (.*?) Between (.*) and (.*) ([A-Z0-9]+)$/).first
      if time1
        expected1 = Time.parse("#{date} #{time1}")
        expected2 = Time.parse("#{date} #{time2}")
        make_org_entry "usps delivery of #{tracking}", 'usps:@waiting', '#C',
                       "<#{expected1.strftime('%F %a %H:%M')}>--<#{expected2.strftime('%F %a %H:%M')}>",
                       "https://tools.usps.com/go/TrackConfirmAction?tLabels=#{tracking}\n" \
                       "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
      else
        _by, date, tracking = headers['Subject'].scan(/Delivery (on|by) (.*?) ([A-Z0-9]+)$/).first
        expected = Time.parse(date)
        make_org_entry "usps delivery of #{tracking}", 'usps:@waiting', '#C',
                       "<#{expected.strftime('%F %a %H:%M')}>",
                       "https://tools.usps.com/go/TrackConfirmAction?tLabels=#{tracking}\n" \
                       "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
      end
    end
  end
end

class MH_AmazonVideoOrder < MessageHandler
  def self.match(headers)
    headers['Subject']&.include?('Amazon.com order of') &&
      headers['From'] == '"Amazon.com" <digital-no-reply@amazon.com>'
  end

  def handle(message, headers)
    payload = (gmail.get_user_message 'me', message.id).payload
    body = payload.parts[0].body.data
    order = headers['Subject'][/Amazon.com order of (.*)\./, 1]
    total = body[/Grand Total:\s+(\$.*)\r\n/, 1]
    $logger.info "#{order} #{total}"

    detail = total.to_s
    make_org_entry "order of #{order}", 'amazon:@quicken', '#C',
                   "<#{Time.now.strftime('%F %a')}>",
                   "#{detail}\nhttps://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end

class GMailTender < Thor
  no_commands do
    def redirect_output
      unless LOGFILE == 'STDOUT'
        logfile = File.expand_path(LOGFILE)
        FileUtils.mkdir_p(File.dirname(logfile), mode: 0o755)
        FileUtils.touch logfile
        File.chmod 0o644, logfile
        $stdout.reopen logfile, 'a'
      end
      $stderr.reopen $stdout
      $stdout.sync = $stderr.sync = true
    end

    def setup_logger
      redirect_output if options[:log]

      $logger = Logger.new $stdout
      $logger.level = options[:verbose] ? Logger::DEBUG : Logger::INFO
      $logger.info 'starting'
    end
  end

  class_option :log,     type: :boolean, default: true, desc: "log output to #{LOGFILE}"
  class_option :verbose, type: :boolean, aliases: '-v', desc: 'increase verbosity'

  desc 'auth', 'Authorize the application with google services'
  def auth
    setup_logger
    #
    # initialize the API
    #
    begin
      service = Google::Apis::GmailV1::GmailService.new
      service.client_options.application_name = APPLICATION_NAME
      service.authorization = authorize !options[:log]
      @gmail = service

      service = Google::Apis::CalendarV3::CalendarService.new
      service.client_options.application_name = APPLICATION_NAME
      service.authorization = authorize !options[:log]
      @gcal = service
    rescue StandardError => e
      $logger.error e.message
      $logger.error e.backtrace.inspect
    end
  end

  desc 'scan', 'Scan emails'
  method_option :dry_run, type: :boolean, aliases: '-n', desc: "don't create tasks or change gmail"
  def scan
    auth

    begin
      #
      # scan unread inbox messages
      #
      messages = (@gmail.list_user_messages 'me', q: 'in:inbox is:unread').messages

      $logger.info "#{messages.nil? ? 'no' : messages.length} unread messages found in inbox"

      messages&.each do |message|
        $logger.debug "- #{message.id}"

        headers = {}
        content = (@gmail.get_user_message 'me', message.id).payload
        content.headers.each do |header|
          $logger.debug "#{header.name} => #{header.value}"
          headers[header.name] = header.value
        end

        MessageHandler.dispatch @gcal, @gmail, options, message, headers
      end

      #
      # scan context folders
      #
      ['@agendas', '@calls', '@errands', '@home', '@quicken', '@view', '@waiting', '@work'].each do |context|
        threads = (@gmail.list_user_threads 'me', q: "in:#{context}").threads

        $logger.info "#{threads.nil? ? 'no' : threads.length} messages found in #{context}"

        threads&.each do |thread|
          # find most recent message in thread list
          messages = (@gmail.get_user_thread 'me', thread.id, fields: 'messages(id,internalDate)').messages
          message = messages.max_by(&:internal_date)
          $logger.debug "- #{message.id}"

          headers = {}
          content = (@gmail.get_user_message 'me', message.id).payload
          content.headers.each do |header|
            $logger.debug "#{header.name} => #{header.value}"
            headers[header.name] = header.value
          end

          MessageHandler.refile @gmail, options, context, thread, message, headers
        end
      end
    rescue StandardError => e
      $logger.error e.message
      $logger.error e.backtrace.inspect
    end

    $logger.info 'done'
  end
end

GMailTender.start
