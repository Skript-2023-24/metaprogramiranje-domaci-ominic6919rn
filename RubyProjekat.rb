require "google_drive"
session = GoogleDrive::Session.from_config("config.json")

session.files.each do |file|
  p file.title
end
