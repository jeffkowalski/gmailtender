#!/usr/bin/env ruby

# GMail API: https://developers.google.com/gmail/api/quickstart/ruby
#            https://developers.google.com/gmail/api/v1/reference/
#            https://console.developers.google.com
# Google API Ruby Client:  https://github.com/google/google-api-ruby-client

require 'googleauth'
require 'googleauth/stores/file_token_store'
require 'google/apis/gmail_v1'
require 'google/apis/calendar_v3'

require 'andand'
require 'fileutils'
require 'logger'
require "net/http"
require "uri"
require 'base64'
require 'thor'
require 'resolv-replace'


LOGFILE = File.join(Dir.home, '.gmailtender.log')

OOB_URI = 'urn:ietf:wg:oauth:2.0:oob'
APPLICATION_NAME = 'gmailtender'
CLIENT_SECRETS_PATH = 'client_secret.json'
CREDENTIALS_PATH = File.join(Dir.home, '.credentials', "gmailtender.yaml")
SCOPE = [Google::Apis::GmailV1::AUTH_GMAIL_MODIFY, 'https://www.googleapis.com/auth/calendar']

##
# Ensure valid credentials, either by restoring from the saved credentials
# files or intitiating an OAuth2 authorization. If authorization is required,
# the user's default browser will be launched to approve the request.
#
# @return [Google::Auth::UserRefreshCredentials] OAuth2 credentials
def authorize interactive
  FileUtils.mkdir_p(File.dirname(CREDENTIALS_PATH))
  client_id = Google::Auth::ClientId.from_file(CLIENT_SECRETS_PATH)
  token_store = Google::Auth::Stores::FileTokenStore.new(file: CREDENTIALS_PATH)
  authorizer = Google::Auth::UserAuthorizer.new(client_id, SCOPE, token_store)
  user_id = 'default'
  credentials = authorizer.get_credentials(user_id)
  if credentials.nil? and interactive
    url = authorizer.get_authorization_url(base_url: OOB_URI)
    code = ask("Open the following URL in the browser and enter the resulting code after authorization\n" + url)
    credentials = authorizer.get_and_store_credentials_from_code(
      user_id: user_id, code: code, base_url: OOB_URI)
  end
  credentials
end


class MessageHandler
  attr_accessor :gmail, :gcal

  def initialize params = {}
    params.each { |key, value| send "#{key}=", value }
  end


  def self.descendants
    ObjectSpace.each_object(Class).select { |klass| klass < self }
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
    if (response.code != '200')
      $logger.error "make_org_entry gave response #{response.code} #{response.message}"
    end
    return response.code == '200'
  end


  def archive message, label_name='INBOX'
    $logger.info "archiving #{message.id}"

    labels = (gmail.list_user_labels 'me').labels
    label_id = labels.select {|label| label.name==label_name }.first.id

    mmr = Google::Apis::GmailV1::ModifyMessageRequest.new(remove_label_ids: [label_id])
    gmail.modify_message 'me', message.id, mmr
  end


  def unlabel_thread thread, label_name
    $logger.info "archiving #{thread.id}"

    labels = (gmail.list_user_labels 'me').labels
    label_id = labels.select {|label| label.name==label_name }.first.id

    mmr = Google::Apis::GmailV1::ModifyThreadRequest.new(remove_label_ids: [label_id])
    gmail.modify_thread 'me', thread.id, mmr
  end


  def self.dispatch gcal, gmail, message, headers
    $logger.info headers['Subject']
    $logger.info headers['From']
    MessageHandler.descendants.each do |handler|
      $logger.debug "matching #{handler}"
      begin
        if handler.match headers
          handler.new(:gcal => gcal, :gmail => gmail).process message, headers
          break
        end
      rescue Exception => e
        $logger.error e.message
        $logger.error e.backtrace.inspect
      end
    end
  end


  def process message, headers
    $logger.info "(#{self.class})"

    response = handle message, headers

    if (response)
      archive message
    end
  end


  def self.friendly_name addr
    name = addr[/"?(.*?)"?\s</, 1] || addr[/<?(.*?)@/, 1]
    if not name.nil?
      name.downcase!
      name.gsub!(/[\. ]/, '_')
    end
    return name
  end


  def self.refile gmail, context, thread, message, headers
    $logger.info headers['Subject']
    $logger.info headers['From']

    full_context = context
    if context == '@waiting'
      to = friendly_name headers['To']
      if to == 'jeff_kowalski'
        to = friendly_name headers['From']
      end
      full_context = [to, context].join(':')
    end

    handler = MessageHandler.new(:gmail => gmail)
    response = handler.make_org_entry headers['Subject'], full_context, '#C',
                                      "<#{Time.now.strftime('%F %a')}>",
                                      "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
    if (response)
      handler.unlabel_thread thread, context
    end
  end
end


class MH_AllstateBill < MessageHandler
  def self.match headers
    headers['Subject'] == 'Allstate: Your billing document is ready to view online' &&
      headers['From'] == 'Allstate My Account <allstate@trns01.allstate-email.com>'
  end

  def handle message, headers
    raw = (gmail.get_user_message 'me', message.id, format: 'raw').raw
    # Policy Number: XXXX63191
    # Policy Type: Auto - private passenger voluntary
    # Document Title: Auto policy schedule for the Recurring Credit Card Pay Plan
    # Due Date: Payments scheduled for the 16th
    # Minimum Amount Due: See Schedule
    detail = raw[/(Policy Number:.*?Minimum Amount Due:.*?\n)/m, 1]
    detail.gsub!("\015", '')
    puts detail
    return make_org_entry 'allstate bill available', '@quicken', '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          detail +
                          "https://myaccount.allstate.com/anon/login/login.aspx\n" +
                          "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end


class MH_ChaseCreditCardStatement < MessageHandler
  def self.match headers
    headers['Subject'] == 'Your credit card statement is available online' &&
      headers['From'] == 'Chase <no-reply@alertsp.chase.com>'
  end

  def handle message, headers
    return make_org_entry 'chase credit card statement available', 'chase:@quicken', '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          "https://stmts.chase.com/stmtslist?AI=16258879\n" +
                          "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end


class MH_ChaseMortgageStatement < MessageHandler
  def self.match headers
    headers['Subject'] == 'Your mortgage statement is available online.' &&
      headers['From'] == 'Chase <no-reply@alertsp.chase.com>'
  end

  def handle message, headers
    return make_org_entry 'chase mortgage statement available', 'chase:@quicken', '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          "https://stmts.chase.com/stmtslist?AI=475283320\n" +
                          "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end


class MH_CapitalOneTransfer < MessageHandler
  def self.match headers
    headers['Subject'] == "Transfer Money Notice" &&
      headers['From'].include?("capitalone.com")
  end

  def handle message, headers
    raw = (gmail.get_user_message 'me', message.id, format: 'raw').raw
    # e.g:
    #  Amount: $39.99
    #  From: Orange Parker Allowance, XXXXXX1099
    #  To: Orange Checking, XXXXXX6515
    #  Memo: game
    #  Transferred On: 08/22/2015
    detail = raw[/(Amount:.*?Transferred On:.*?\n)/m, 1]
    detail.gsub!("\015", '')
    return make_org_entry 'capital one transfer money notice', 'capitalone:@quicken', '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          detail + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end


class MH_CapitalOneStatement < MessageHandler
  def self.match headers
    headers['Subject']&.include?("eStatement's now available") &&
      headers['From'] == "Capital One <capitalone@email.capitalone.com>"
  end

  def handle message, headers
    return make_org_entry 'account statement available', 'capitalone:@quicken', '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          "https://secure.capitalone360.com/myaccount/banking/login.vm\n" +
                          "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end


class MH_PaypalStatement < MessageHandler
  def self.match headers
    headers['Subject']&.include?("account statement is available") &&
      headers['From'] == 'PayPal Statements <paypal@e.paypal.com>'
  end

  def handle message, headers
    payload = (gmail.get_user_message 'me', message.id).payload
    body = payload.body.data
    detail = '' + body[/<a.*?href="(.*?)".*?View Statement<\/a>/m, 1] + "\n"
    return make_org_entry 'account statement available', 'paypal:@quicken', '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          detail + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end


class MH_SquareReceipt < MessageHandler
  def self.match headers
    headers['Subject']&.include?("Receipt from") &&
      headers['From']&.include?("via Square <receipts@messaging.squareup.com>")
  end

  def handle message, headers
    payload = (gmail.get_user_message 'me', message.id).payload
    body = payload.parts[0].body.data;

    # You paid $10.00 with your AMEX ending in 2008 to Pizza Politana on Aug 22

    amount = '' + body[/You paid (\$\d+\.\d+) with/, 1]
    account = '' + body[/with your (\S+) ending in/, 1].downcase
    payee = '' + body[/to (.*?) on/, 1].downcase
    return make_org_entry "receipt for #{payee}", "#{account}:@quicken", '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          amount + "\n" + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end


class MH_CloverReceipt < MessageHandler
  def self.match headers
    headers['Subject']&.include?("Your receipt from") &&
      headers['From']&.include?("<app@clover.com>")
  end

  def handle message, headers
    payload = (gmail.get_user_message 'me', message.id).payload
    body = payload.parts[0].body.data;

    # $6.95 total

    amount = '' + body[/(\$\d+\.\d+) total/, 1]
    payee = '' +  headers['Subject'][/Your receipt from (.*)/, 1].downcase
    return make_org_entry "receipt for #{payee}", "@quicken", '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          amount + "\n" + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end


class MH_SquareReceipt < MessageHandler
  def self.match headers
    headers['Subject']&.include?("Receipt from") &&
      headers['From']&.include?("via Square <receipts@messaging.squareup.com>")
  end

  def handle message, headers
    payload = (gmail.get_user_message 'me', message.id).payload
    body = payload.parts[0].body.data;

    # You paid $10.00 with your AMEX ending in 2008 to Pizza Politana on Aug 22

    amount = '' + body[/You paid (\$\d+\.\d+) with/, 1]
    account = '' + body[/with your (\S+) ending in/, 1].downcase
    payee = '' + body[/to (.*?) on/, 1].downcase
    return make_org_entry "receipt for #{payee}", "#{account}:@quicken", '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          amount + "\n" + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end


class MH_PershingStatement < MessageHandler
  def self.match headers
    headers['Subject'] =='Account Statement Notification' &&
      headers['From'] == 'pershing@advisor.netxinvestor.com'
  end

  def handle message, headers
    return make_org_entry 'account statement available', 'pershing:@quicken', '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end


class MH_PGEStatement < MessageHandler
  def self.match headers
    headers['Subject'] =='Your PG&E Energy Statement is Ready to View' &&
      headers['From'] == 'CustomerServiceOnline@billpay.pge.com'
  end

  def handle message, headers
    return make_org_entry 'pg&e statement available', 'amex:@quicken', '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          "http://www.pge.com/MyEnergy\nhttps://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end


class MH_ATT_Wireless_Bill < MessageHandler
  def self.match headers
    headers['Subject'] =='Your AT&T wireless bill is ready to view'
  end

  def handle message, headers
    return make_org_entry 'at&t wireless bill ready', 'amex:@quicken', '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end


class MH_PeetsReload < MessageHandler
  def self.match headers
    headers['Subject']&.include?("Your Peet's Card Reload Order") &&
      headers['From'].include?("<customerservice@peets.com>")
  end

  def handle message, headers
    return make_org_entry 'peet\'s card reload order', 'amex:@quicken', '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          "$50\n" +
                          "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end


class MH_ComcastBill < MessageHandler
  def self.match headers
    headers['Subject'] == 'Your bill is ready' &&
      headers['From'] == 'XFINITY My Account <online.communications@alerts.comcast.net>'
  end
  def handle message, headers
    return make_org_entry 'comcast bill ready', 'amex:@quicken', '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          "https://customer.xfinity.com/Secure/MyAccount/\n" +
                          "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end


class MH_EtradeStatement < MessageHandler
  def self.match headers
    headers['Subject'] == 'You have a new account statement from E*TRADE Securities' &&
      headers['From'] == '"E*TRADE SECURITIES LLC" <etrade_stmt_mbox@statement.etradefinancial.com>'
  end

  def handle message, headers
    return make_org_entry 'etrade statement available', 'etrade:@quicken', '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          "https://edoc.etrade.com/e/t/onlinedocs/docsearch?doc_type=stmt\n" +
                          "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end


class MH_AmericanExpressStatement < MessageHandler
  def self.match headers
    headers['Subject']&.index(/Your .* Statement is Ready/) &&
      headers['From'] == 'American Express <AmericanExpress@welcome.aexp.com>'
  end

  def handle message, headers
    return make_org_entry 'account statement available', 'amex:@quicken', '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          "https://online.americanexpress.com/myca/statementimage/us/welcome.do?request_type=authreg_StatementCycles&Face=en_US&sorted_index=0\n" +
                          "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end


class MH_VerizonBill < MessageHandler
  def self.match headers
    headers['Subject'] == 'Your online bill is available.' &&
      headers['From'] == 'Verizon Wireless <VZWMail@ecrmemail.verizonwireless.com>'
  end

  def handle message, headers
    return make_org_entry 'verizon bill available', 'amex:@quicken', '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          "https://ebillpay.verizonwireless.com/vzw/accountholder/mybill/BillingSummary.action\n" +
                          "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end


class MH_AmazonSubscribeAndSave < MessageHandler
  def self.match headers
    headers['Subject'] == 'Amazon Subscribe & Save: Review Your Monthly Delivery' &&
      headers['From'] == '"Amazon Subscribe & Save" <no-reply@amazon.com>'
  end

  def handle message, headers
    return make_org_entry 'review amazon subscribe and save delivery', 'amazon:@home', '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          "https://www.amazon.com/manageyoursubscription\n" +
                          "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end


class MH_AmazonOrder < MessageHandler
  def self.match headers
    headers['Subject']&.include?('Your Amazon.com order') &&
      headers['From'] == '"Amazon.com" <auto-confirm@amazon.com>'
  end

  def handle message, headers
    payload = (gmail.get_user_message 'me', message.id).payload
    body = payload.parts[0].body.data;
    order = headers['Subject'][/Your Amazon.com order of (.*)\./, 1] ||
            body[/You ordered\s+(".*?")\s*\.\r\n/m, 1]
    url = body[/View or manage your orders in Your Orders:\r\n?(https:.*?)\r\n/m, 1]
    delivery = body[/\s*((:?(:?Guaranteed|Estimated) delivery date)|(:?Arriving)):\s*\r\n\s*(?<date>.*?)\r\n/m, :date]
    delivery = (delivery.nil? ? Time.now : Date.parse(delivery)).strftime('%F %a')
    total = body[/Order Total: (\$.*)\r\n/, 1]
    $logger.info "#{order} #{url} #{delivery} #{total}"

    detail = "#{url}\n#{total}"
    response = make_org_entry "order of #{order}", 'amazon:@quicken', '#C',
                              "<#{Time.now.strftime('%F %a')}>",
                              detail + "\n" + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
    if (response)
      detail = url
      response = make_org_entry "delivery of #{order}", 'amazon:@waiting', '#C',
                                "<#{delivery}>",
                                detail + "\n" + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
    end
    return response
  end
end


class MH_AmazonVideoOrder < MessageHandler
  def self.match headers
    headers['Subject']&.include?('Amazon.com order of') &&
      headers['From'] == '"Amazon.com" <digital-no-reply@amazon.com>'
  end

  def handle message, headers
    payload = (gmail.get_user_message 'me', message.id).payload
    body = payload.parts[0].body.data;
    order = headers['Subject'][/Amazon.com order of (.*)\./, 1]
    total = body[/Grand Total:\s+(\$.*)\r\n/, 1]
    $logger.info "#{order} #{total}"

    detail = "#{total}"
    return make_org_entry "order of #{order}", 'amazon:@quicken', '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          detail + "\n" + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end


class MH_VanguardStatement < MessageHandler
  def self.match headers
    headers['Subject'] == 'Your Vanguard statement is ready' &&
      headers['From'] == 'Vanguard <ParticipantServices@vanguard.com>'
  end

  def handle message, headers
    return make_org_entry 'account statement available', 'vanguard:@quicken', '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          "https://retirementplans.vanguard.com/VGApp/pe/PublicHome\n" +
                          "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end


class MH_WorkdayFeedbackRequest < MessageHandler
  def self.match headers
    headers['Subject'] == 'Feedback is requested' &&
      headers['From'] == 'AutoNotification workday <autodesk@myworkday.com>'
  end

  def handle message, headers
    #Mark Davis (110932) has requested that you provide feedback on Anthony Ruto
    payload = (gmail.get_user_message 'me', message.id).payload
    body = payload.parts[0].parts[0].body.data;
    requester, employee = body.match(/<span>([^<].*?) \(\d+\) has requested that you provide feedback on (.*?) - Please visit your Workday inbox/).captures
    detail = body[/<a href="(https:\/\/.*?)">Click Here to view the notification details/, 1]
    return make_org_entry "provide feedback on #{employee} to #{requester}", '@work', '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          detail + "\n" + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  end
end


class MH_SelectASpot < MessageHandler
  def self.match headers
    headers['Subject'] == 'Select-a-Spot.com: Reservation Successful' &&
      headers['From'] == 'Select-a-Spot <notifications@select-a-spot.com>'
  end

  def handle message, headers
    payload = (gmail.get_user_message 'me', message.id).payload
    body = payload.parts[0].body.data;
    #  <strong>Effective Dates:</strong> April 3, 2017 - April 3, 2017<br/>
    #  <strong>Facility Details:</strong> Orinda Station, 11 Camino Pablo, Orinda, CA 94563, <a href="http://www.bart.gov/sites/default/files/docs/BART_Parking_Orinda.pdf">view facility map</a><br/>
    #  <strong>Permit Number: </strong>1260089<br/>
    #  <p>Price: $6.00</p>
    dates = body.match(/<strong>Effective Dates:<\/strong> (.*) - (.*)<br\/>/).captures
    facility = body[/<strong>Facility Details:<\/strong> (.*), <a href=".*">view facility map<\/a><br\/>/, 1]
    station = facility[/(.*?),/, 1].downcase
    permit = body[/<strong>Permit Number: <\/strong>(.*)<br\/>/, 1]
    price = body[/<p>Price: \$(.*)<\/p>/, 1]
    $logger.info "#{dates} #{facility} #{permit} #{price}"

    (Date.parse(dates[0])..Date.parse(dates[1])).each do |date|

      old_events = gcal.list_events('primary',
                                    single_events: true,
                                    q: 'bart parking',
                                    time_min: Time.parse(date.to_s + ' 00:00:00 Pacific Time').iso8601,
                                    time_max: Time.parse(date.to_s + ' 23:59:59 Pacific Time').iso8601,
                                    fields: 'items(id,summary,location,organizer,attendees,description,start,end),next_page_token')
      old_events.items.each do |event|
        $logger.info "deleting #{event.summary} #{event.id}"
        gcal.delete_event 'primary', event.id
      end

      new_event = Google::Apis::CalendarV3::Event.new(
        {
          summary: "bart parking at #{station}",
          location: facility,
          description: "permit #{permit}",
          start: {
            date_time: Time.parse(date.to_s + ' 09:30:00 Pacific Time').iso8601,
            # 2015-05-28T09:00:00-07:00
            # time_zone: 'America/Los_Angeles'
          },
          end: {
            date_time: Time.parse(date.to_s + ' 10:00:00 Pacific Time').iso8601,
            # time_zone: 'America/Los_Angeles'
          }
        }
      )

      $logger.info new_event

      # https://developers.google.com/google-apps/calendar/v3/reference/events/insert
      result = gcal.insert_event 'primary', new_event
    end

    return make_org_entry "select-a-spot #{price}", 'amazon_visa:@quicken', '#C',
                          "<#{Time.now.strftime('%F %a')}>",
                          dates[0] + "\n" + station + "\n" + permit

  end
end


class GMailTender < Thor
  no_commands {
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

    def setup_logger
      redirect_output if options[:log]

      $logger = Logger.new STDOUT
      $logger.level = options[:verbose] ? Logger::DEBUG : Logger::INFO
      $logger.info 'starting'
    end
  }

  class_option :log,     :type => :boolean, :default => true, :desc => "log output to ~/.gmailtender.log"
  class_option :verbose, :type => :boolean, :aliases => "-v", :desc => "increase verbosity"

  desc "auth", "Authorize the application with google services"
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
    rescue Exception => e
      $logger.error e.message
      $logger.error e.backtrace.inspect
    end
  end

  desc "scan", "Scan emails"
  def scan
    auth

    begin

      #
      # scan unread inbox messages
      #
      messages = (@gmail.list_user_messages 'me', q: 'in:inbox is:unread').messages

      $logger.info "#{messages.nil? ? 'no' : messages.length} unread messages found in inbox"

      messages.andand.each do |message|
        $logger.debug "- #{message.id}"

        headers = {}
        content = (@gmail.get_user_message 'me', message.id).payload
        content.headers.each do |header|
          $logger.debug "#{header.name} => #{header.value}"
          headers[header.name] = header.value
        end

        MessageHandler.dispatch @gcal, @gmail, message, headers
      end


      #
      # scan context folders
      #
      ['@agendas', '@calls', '@errands', '@home', '@quicken', '@view', '@waiting', '@work'].each do |context|
        threads = (@gmail.list_user_threads 'me', q: "in:#{context}").threads

        $logger.info "#{threads.nil? ? 'no' : threads.length} messages found in #{context}"

        threads.andand.each do |thread|
          # find most recent message in thread list
          messages = (@gmail.get_user_thread 'me', thread.id, fields: "messages(id,internalDate)").messages
          message = messages.sort_by(&:internal_date).last
          $logger.debug "- #{message.id}"

          headers = {}
          content = (@gmail.get_user_message 'me', message.id).payload
          content.headers.each do |header|
            $logger.debug "#{header.name} => #{header.value}"
            headers[header.name] = header.value
          end

          MessageHandler.refile @gmail, context, thread, message, headers
        end
      end

    rescue Exception => e
      $logger.error e.message
      $logger.error e.backtrace.inspect
    end

    $logger.info 'done'
  end
end

GMailTender.start
