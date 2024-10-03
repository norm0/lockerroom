require 'net/http'
require 'icalendar'
require 'uri'
require 'csv'

# iCal feed URL (using https)
ical_feed_url = 'https://www.armstrongcooperhockey.org/ical_feed?tags=8603019'

# Family names provided
family_names = %w[
  Becker Hastings Opel Gorgos Larsen
  Anderson Orstad Campos Powell Tousignant Marshall Johnson Wulff
]

# Fetch the iCal feed using Net::HTTP
uri = URI(ical_feed_url)
response = Net::HTTP.get(uri)

# Parse the iCal feed
calendar = Icalendar::Calendar.parse(response).first

# Filter for events that have a start and end time and are at specific locations
filtered_events = calendar.events.select do |event|
  event.dtstart && event.dtend &&
    ['New Hope North', 'New Hope South', 'Breck'].include?(event.location)
end

# Helper method to format dates with abbreviated day and month names
def format_friendly_date(datetime)
  datetime.strftime('%a, %b %-d, %Y at %-I:%M %p')
end

# Initialize a new iCal feed for locke