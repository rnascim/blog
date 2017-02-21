# require 'java'
# import java.util.Properties
# import com.sap.conn.jco.JCoDestination
# import com.sap.conn.jco.JCoDestinationManager
# import com.sap.conn.jco.ext.DestinationDataProvider;

module SapConnHelper
	# class SAPConnection

	# 	def self.instance
	# 		@_instance ||= new
	# 	end
		
	# 	attr_reader :property, :conn

	# 	def initialize
	# 		prop = java.util.Properties.new
	# 		Settings.each do |k, v|
	# 			prop.setProperty(k.to_s, v.to_s)
	# 		end			

	# 		dest = JCoDestinationManager::getDestination prop
	# 		logger.debug("Destination attributes: " + dest.getAttributes)

	# 		dataProvider = com.sap.conn.jco.ext.DestinationDataProvider.new
	# 		dataProvider.changeProperties("nome", prop)
	# 	end

	# end

end
