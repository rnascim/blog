
SapJCo::Configuration.configure ({:default_destination => :test})

rfc =  SapJCo::Function.new(:RFC_SYSTEM_INFO)
out = rfc.execute
puts out[:RFCSI_EXPORT][:RFCHOST2]  # Should print the host name of the SAP application server
