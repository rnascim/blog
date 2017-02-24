# raw_yaml = YAML.load(File.read(File.join(Rails.root, "/config/sidekiq.yml")))
# SIDEKIQ_CONFIG = raw_yaml.merge(raw_yaml[Rails.env])

# Sidekiq::Logging.logger.level = Logger::DEBUG

# Sidekiq.configure_client do |config|
#   config.redis = {
#     :url       => SIDEKIQ_CONFIG[:url],
#     :namespace => SIDEKIQ_CONFIG[:namespace],
#     :size      => SIDEKIQ_CONFIG[:client_connections]
#   }
# end

# Sidekiq.configure_server do |config|
#   config.redis = {
#     :url       => SIDEKIQ_CONFIG[:url],
#     :namespace => SIDEKIQ_CONFIG[:namespace],
#     :size      => SIDEKIQ_CONFIG[:client_connections]
#   }
# end