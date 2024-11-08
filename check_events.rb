require 'net/http'
require 'icalendar'
require 'uri'

# Define teams and configurations
teams = [
  {
    name: '12A',
    ical_feed_url: 'https://www.armstrongcooperhockey.org/ical_feed?tags=8603019',
    family_names: %w[Becker Hastings Opel Gorgos Larsen Anderson Campos Powell Tousignant Marshall Johnson Wulff Orstad
                     Mulcahey],
    spreadsheet_id: ENV['GOOGLE_SHEET_ID_12A']
  },
  {
    name: '12B1',
    ical_feed_url: 'https://www.armstrongcooperhockey.org/ical_feed?tags=8603021',
    family_names: %w[Baer Bimberg Chanthavongsa Hammerstrom Kremer Lane Oas Perpich Ray Reinke Silva-Hammer],
    spreadsheet_id: ENV['GOOGLE_SHEET_ID_12B1']
  },
  {
    name: '10B1',
    ical_feed_url: 'https://www.armstrongcooperhockey.org/ical_feed?tags=8603022',
    family_names: %w[Baer Bowman Hopper Houghtaling Johnson Larsen Markfort Marshall Nanninga Orstad Willey Williamson],
    spreadsheet_id: ENV['GOOGLE_SHEET_ID_10B1']
  },
  {
    name: '10B2',
    ical_feed_url: 'https://www.armstrongcooperhockey.org/ical_feed?tags=8603023',
    family_names: %w[Curry Engholm Froberg Harpel Johnson Mckinnon Oprenchak M-Reberg B-Reberg Sauer Smith Woods],
    spreadsheet_id: ENV['GOOGLE_SHEET_ID_10B2']
  }
]

# Define an exclusion list for locations
exclusion_list = [
  'Skills Off Ice', # Example keywords or patterns
  'Dryland',
  'Goalie Training',
  'Off Ice',
  'Conditioning',
  'Meeting',
  'New Hope North - Skills Off Ice',
  'New Hope Ice Arena, Louisiana Avenue North, New Hope, MN, USA',
  'LRM'
]

# Get the team name from the command-line arguments
team_name_to_process = ARGV[0]

if team_name_to_process.nil?
  puts 'Please provide a team name as a command-line argument.'
  exit
end

# Filter the teams array to include only the specified team
teams_to_process = teams.select { |team| team[:name] == team_name_to_process }

if teams_to_process.empty?
  puts "No team found with the name '#{team_name_to_process}'."
  exit
end

teams_to_process.each do |team|
  puts "Processing team: #{team[:name]}"

  # Fetch the iCal feed using Net::HTTP
  uri = URI(team[:ical_feed_url])
  response = Net::HTTP.get(uri)

  # Parse the iCal feed
  calendar = Icalendar::Calendar.parse(response).first

  # Filter out events that match the exclusion list
  filtered_events = calendar.events.reject do |event|
    event.dtstart.nil? || event.dtend.nil? || exclusion_list.any? { |term| event.location&.include?(term) }
  end

  # Print the filtered events
  filtered_events.each do |event|
    puts "Event: #{event.summary}"
    puts "Location: #{event.location}"
    puts "Start Date: #{event.dtstart}"
    puts "End Date: #{event.dtend}"
    puts '-' * 40
  end
end
