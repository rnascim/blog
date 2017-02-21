# Load the Rails application.
require_relative 'application'

# Initialize the Rails application.
Rails.application.initialize!

$CLASSPATH << 'lib'
require 'sapjco3.jar'