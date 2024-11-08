require 'net/http'
require 'icalendar'
require 'uri'
require 'csv'
require 'google/apis/sheets_v4'
require 'googleauth'
require 'googleauth/stores/file_token_store'
require 'json'
require 'stringio'
require 'active_support/time'

APPLICATION_NAME = 'Google Sheets API Ruby Integration'
SCOPE = ['https://www.googleapis.com/auth/spreadsheets']
CREDENTIALS_PATH = 'credentials.json'
TOKEN_PATH = 'token.yaml'

# Define the authorize method to handle Google OAuth2
def authorize
  unless File.exist?(CREDENTIALS_PATH)
    raise "Missing #{CREDENTIALS_PATH}. Please download it from the Google Cloud Console."
  end

  client_id = Google::Auth::ClientId.from_file(CREDENTIALS_PATH)
  token_store = Google::Auth::Stores::FileTokenStore.new(file: TOKEN_PATH)
  authorizer = Google::Auth::UserAuthorizer.new(client_id, SCOPE, token_store)
  user_id = 'default'
  credentials = authorizer.get_credentials(user_id)
  if credentials.nil?
    url = authorizer.get_authorization_url(base_url: OOB_URI)
    puts "Open the following URL in your browser and authorize the application:\n#{url}"
    print 'Enter the authorization code: '
    code = gets.chomp
    credentials = authorizer.get_and_store_credentials_from_code(user_id:, code:, base_url: OOB_URI)
  end
  credentials
end

# Initialize the Google Sheets API service
def initialize_service
  service = Google::Apis::SheetsV4::SheetsService.new
  service.client_options.application_name = APPLICATION_NAME
  service.authorization = authorize
  service
end

# Method to write team data to individual sheets
def write_team_data_to_individual_sheets(service, team, data)
  headers = ['Event', 'Location', 'Date', 'Time', 'Duration (minutes)', 'Locker Room Monitor']
  values = [headers] + # Replace nils with empty strings
           data.map do |row|
                                    row.map do |cell|
                                      cell.nil? ? '' : cell.to_s
                                    end
           end
  range = 'Sheet1!A1:F'
  value_range = Google::Apis::SheetsV4::ValueRange.new(values:)
  service.update_spreadsheet_value(team[:spreadsheet_id], range, value_range, value_input_option: 'RAW')
end

# Team configurations for 12A, 12B1, 10B1, and 10B2
teams = [
  {
    name: '12A',
    ical_feed_url: 'https://www.armstrongcooperhockey.org/ical_feed?tags=8603019',
    family_names: %w[Becker Hastings Opel Gorgos Larsen Anderson Campos Powell Tousignant Marshall Johnson Wulff Orstad
                     Mulcahey],
    spreadsheet_id: 'your_spreadsheet_id_12A' # Replace with your actual spreadsheet ID
  },
  {
    name: '12B1',
    ical_feed_url: 'https://www.armstrongcooperhockey.org/ical_feed?tags=8603021',
    family_names: %w[Baer Bimberg Chanthavongsa Hammerstrom Kremer Lane Oas Perpich Ray Reinke Silva-Hammer],
    spreadsheet_id: 'your_spreadsheet_id_12B1' # Replace with your actual spreadsheet ID
  },
  {
    name: '10B1',
    ical_feed_url: 'https://www.armstrongcooperhockey.org/ical_feed?tags=8603022',
    family_names: %w[Baer Bowman Hopper Houghtaling Johnson Larsen Markfort Marshall Nanninga Orstad Willey Williamson],
    spreadsheet_id: 'your_spreadsheet_id_10B1' # Replace with your actual spreadsheet ID
  },
  {
    name: '10B2',
    ical_feed_url: 'https://www.armstrongcooperhockey.org/ical_feed?tags=8603023',
    family_names: %w[Curry Engholm Froberg Harpel Johnson Mckinnon Oprenchak M-Reberg B-Reberg Sauer Smith Woods],
    spreadsheet_id: 'your_spreadsheet_id_10B2' # Replace with your actual spreadsheet ID
  }
]

# Locations that require locker room monitors
locations_with_monitors = ['New Hope North', 'New Hope South', 'Breck', 'Orono Ice Arena (ag)', 'Northeast (ag)',
                           'SLP East (ag)', 'MG West (ag)', 'PIC A (ag)', 'PIC C (ag)', 'Hopkins Pavilion (ag)', 'Thaler (ag)', 'SLP West (ag)', 'Delano Arena']

# Track the last assigned family for each team to avoid back-to-back assignments
last_assigned_family = {}

# Initialize the Google Sheets API service
service = initialize_service

teams.each do |team|
  # Initialize the assignment count for each family in the team if not already present
  team[:family_names].each { |family| assignment_counts[family] ||= 0 }

  # Fetch the iCal feed using Net::HTTP for the current team
  uri = URI(team[:ical_feed_url])
  response = Net::HTTP.get(uri)

  # Parse the iCal feed
  calendar = Icalendar::Calendar.parse(response).first

  # Initialize a new iCal feed for locker room monitor events
  lrm_calendar = Icalendar::Calendar.new

  # Assign a locker room monitor for each event, cycling through family names
  csv_data = calendar.events.each_with_index.map do |event, _index|
    # Generate a unique event ID (e.g., using the event UID)
    event_id = event.uid

    # Skip events without a start or end time
    next if event.dtstart.nil? || event.dtend.nil?

    # Skip all-day events with "LRM" in the summary or description
    if event.dtstart.is_a?(Icalendar::Values::Date) && (event.summary&.include?('LRM') || event.description&.include?('LRM'))
      next
    end

    # Convert Icalendar::Values::DateTime to Ruby Time object
    start_time = event.dtstart.to_time
    end_time = event.dtend.to_time

    # Format the date and time separately (human-readable)
    raw_date = start_time.strftime('%Y-%m-%d')
    formatted_date = start_time.strftime('%a %I:%M %p').downcase
    duration_in_minutes = ((end_time - start_time) / 60).to_i

    # Balanced random assignment of locker room monitor per team
    locker_room_monitor = assigned_events[event_id] || begin
      shuffled_families = team[:family_names].shuffle
      eligible_families = shuffled_families.reject { |family| family == last_assigned_family[team[:name]] }
      family_with_fewest_assignments = eligible_families.min_by { |family| assignment_counts[family] }
      assignment_counts[family_with_fewest_assignments] += 1
      assigned_events[event_id] = family_with_fewest_assignments
      last_assigned_family[team[:name]] = family_with_fewest_assignments
      family_with_fewest_assignments
    end

    # Prepare data for Google Sheets
    [event.summary, event.location, raw_date, formatted_date, duration_in_minutes, locker_room_monitor]
  end.compact

  # Clear existing data and write new data to Google Sheets
  clear_google_sheet_data(service, team[:spreadsheet_id], 'Sheet1!A1:F')
  write_team_data_to_individual_sheets(service, team, csv_data)

  # Save assignment counts to ensure persistence
  CSV.open("#{@assignment_counts_file}_#{team[:name]}.csv", 'w') do |csv|
    csv << %w[Family Count]
    assignment_counts.each do |family, count|
      csv << [family, count]
    end
  end

  # Save assigned events to ensure persistence
  CSV.open("#{@assigned_events_file}_#{team[:name]}.csv", 'w') do |csv|
    csv << %w[EventID Locker_Room_Monitor]
    assigned_events.each do |event_id, monitor|
      csv << [event_id, monitor]
    end
  end
end

# Display the family counts by team
puts "\nLocker Room Monitor Assignment Counts by Team:"
teams.each do |team|
  puts "\nTeam #{team[:name]}:"
  team[:family_names].each do |family|
    puts "#{family}: #{assignment_counts[family]}"
  end
end
