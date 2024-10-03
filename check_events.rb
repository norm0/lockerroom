require 'net/http'
require 'icalendar'
require 'uri'

# iCal feed URL (using https)
ical_feed_url = 'https://www.armstrongcooperhockey.org/ical_feed?tags=8603019'

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

# Print the filtered events
filtered_events.each do |event|
  puts "Event: #{event.summary}"
  puts "Location: #{event.location}"
  puts "Start Date: #{event.dtstart}"
  puts "End Date: #{event.dtend}"
  puts '-' * 40
end
