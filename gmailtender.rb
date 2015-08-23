#!/usr/bin/ruby

require 'google/api_client'
require 'google/api_client/client_secrets'
require 'google/api_client/auth/installed_app'
require 'google/api_client/auth/storage'
require 'google/api_client/auth/storages/file_store'
require 'fileutils'

APPLICATION_NAME = 'Gmail API Quickstart'
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
    puts "Credentials saved to #{CREDENTIALS_PATH}" unless auth.nil?
  end
  auth
end

# Initialize the API
Client = Google::APIClient.new(:application_name => APPLICATION_NAME)
Client.authorization = authorize
Gmail_api = Client.discovered_api('gmail', 'v1')

require "net/http"
require "uri"


def encodeURIcomponent str
  return URI.escape(str, Regexp.new("[^#{URI::PATTERN::UNRESERVED}]"))
end


def make_org_entry heading, context, priority, date, body
  # puts "TODO [#{priority}] #{heading} #{context}"
  # puts "SCHEDULED: #{date}"
  # puts "#{body}"
  title = encodeURIcomponent "[#{priority}] #{heading}  :#{context}:"
  body  = encodeURIcomponent "SCHEDULED: #{date}\n#{body}"
  uri = URI.parse("http://carbon:3333")
  http = Net::HTTP.new(uri.host, uri.port)
  request = Net::HTTP::Get.new("/capture/b/LINK/#{title}/#{body}")
  response = http.request(request)
  return response
end


def archive message
  puts "archiving #{message.id}"
  Client.execute!(
    :api_method => Gmail_api.users.messages.modify,
    :parameters => { 'userId' => 'me', 'id' => message.id },
    :body_object => { 'removeLabelIds' => ['INBOX'] })
end


def process_transfer message, headers
  puts '(process transfer)'
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
  response = make_org_entry 'capital one transfer money notice', '@quicken', '#C', "<#{Time.now.strftime('%F %a')}>", detail + "https://mail.google.com/mail/u/0/#inbox/#{message.id}"
  if (response.code == '200')
    archive message
  end
end


#
def dispatch_message message, headers
  #puts headers['Subject'], headers['From']
  if headers['Subject'] == "Transfer Money Notice" && headers['From'] == 'Capital One 360 <saver@capitalone360.com>'
    process_transfer message, headers
  end
end


# get the user's messages
results = Client.execute!(
  :api_method => Gmail_api.users.messages.list,
  :parameters => { :userId => 'me', :q => "in:inbox" })
#:parameters => { :userId => 'me', :q => "in:inbox is:unread" })

puts "No messages found" if results.data.messages.empty?
results.data.messages.each { |message|
  # See https://developers.google.com/gmail/api/v1/reference/users/messages/get
  content = Client.execute!(
    :api_method => Gmail_api.users.messages.get,
    :parameters => { :userId => 'me', :id => message.id})
  #puts "- #{message.id}"

  # See https://developers.google.com/gmail/api/v1/reference/users/messages#methods
  headers = {}
  content.data.payload.headers.each { |header|
    #puts header.name
    headers[header.name] = header.value
  }

  dispatch_message message, headers
}
