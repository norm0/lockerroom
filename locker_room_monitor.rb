require 'net/http'
require 'icalendar'
require 'uri'
require 'csv'
require 'google/apis/sheets_v4'
require 'googleauth'
require 'json'
require 'stringio'

APPLICATION_NAME = 'Google Sheets API Ruby Integration'
SCOPE = ['https://www.googleapis.com/auth/spreadsheets']

# File to store assignment counts and assigned events
assignment_counts_file = 'assignment_counts.csv'
assigned_events_file = 'assigned_events.csv'

# Google Sheets setup
def setup_google_sheets
  service = Google::Apis::SheetsV4::SheetsService.new
  service.client_options.application_name = APPLICATION_NAME
  service.authorization = authorize
  service
end

# Google Sheets authorization using a service account
def authorize
  # Parse the service account JSON from the GitHub secret
  credentials = JSON.parse(ENV['GOOGLE_SHEETS_CREDENTIALS'])

  # Set up ServiceAccountCredentials using the JSON key
  Google::Auth::ServiceAccountCredentials.make_creds(
    json_key_io: StringIO.new(credentials.to_json),
    scope: SCOPE
  )
end

# Method to write data to Google Sheets
def write_team_data_to_individual_sheets(service, team, data)
  headers = ['Event', 'Location', 'Date', 'Time', 'Duration (minutes)', 'Locker Room Monitor']
  values = [headers] + data
  range = 'Sheet1!A1:F' # Adjust as needed
  value_range = Google::Apis::SheetsV4::ValueRange.new(values:)
  service.update_spreadsheet_value(team[:spreadsheet_id], range, value_range, value_input_option: 'RAW')
end

# Read existing assignment counts from CSV file
assignment_counts = Hash.new(0)
if File.exist?(assignment_counts_file)
  CSV.foreach(assignment_counts_file, headers: true) do |row|
    assignment_counts[row['Family']] = row['Count'].to_i
  end
end

# Read existing assigned events from CSV file
assigned_events = {}
if File.exist?(assigned_events_file)
  CSV.foreach(assigned_events_file, headers: true) do |row|
    assigned_events[row['EventID']] = row['Locker Room Monitor']
  end
end

# Team configurations
teams = [
  {
    name: '12A',
    ical_feed_url: 'https://www.armstrongcooperhockey.org/ical_feed?tags=8603019',
    family_names: %w[Becker Hastings Opel Gorgos Larsen Anderson Campos Powell Tousignant Marshall Johnson Wulff Orstad
                     Mulcahey],
    spreadsheet_id: ENV['GOOGLE_SHEET_ID_12A']
  }
  # Add other team configurations similarly...
]

# Locations that require locker room monitors
locations_with_monitors = ['New Hope North', 'New Hope South', 'Breck', 'Orono Ice Arena (ag)', 'Northeast (ag)',
                           'SLP East (ag)', 'MG West (ag)', 'PIC A (ag)', 'PIC C (ag)', 'Hopkins Pavilion (ag)', 'Thaler (ag)', 'SLP West (ag)', 'Delano Arena']

# Track the last assigned family for each team to avoid back-to-back assignments
last_assigned_family = {}

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
    date_formatted = start_time.strftime('%Y-%m-%d')
    time_formatted = start_time.strftime('%I:%M %p') # 12-hour format with AM/PM

    # Calculate duration in minutes for non-all-day events
    duration_in_minutes = ((end_time - start_time) / 60).to_i

    # Determine if this event location requires a locker room monitor
    locker_room_monitor = if locations_with_monitors.include?(event.location)
                            # Check if the event has already been assigned a monitor
                            assigned_events[event_id] || begin
                              # Shuffle the family names to randomize the order
                              shuffled_families = team[:family_names].shuffle
                              # Select the family with the fewest assignments, avoiding back-to-back assignments
                              eligible_families = shuffled_families.reject do |family|
                                family == last_assigned_family[team[:name]]
                              end
                              family_with_fewest_assignments = eligible_families.min_by do |family|
                                assignment_counts[family]
                              end
                              # Update the count for the selected family
                              assignment_counts[family_with_fewest_assignments] += 1
                              # Track the assignment
                              assigned_events[event_id] = family_with_fewest_assignments
                              # Update the last assigned family for the team
                              last_assigned_family[team[:name]] = family_with_fewest_assignments
                              family_with_fewest_assignments
                            end
                          else
                            nil # No locker room monitor for other locations
                          end

    # Return data for CSV and Google Sheets
    [
      event.summary, event.location, date_formatted, time_formatted, duration_in_minutes, locker_room_monitor || ''
    ]
  end.compact # Remove nil values from the array

  # Write each team's data to its corresponding Google Sheet
  service = setup_google_sheets
  write_team_data_to_individual_sheets(service, team, csv_data)

  # Write updated assignment counts and assigned events to CSV files
  CSV.open(assignment_counts_file, 'w') do |csv|
    csv << %w[Family Count]
    assignment_counts.each { |family, count| csv << [family, count] }
  end

  CSV.open(assigned_events_file, 'w') do |csv|
    csv << ['EventID', 'Locker Room Monitor']
    assigned_events.each { |event_id, monitor| csv << [event_id, monitor] }
  end
end
