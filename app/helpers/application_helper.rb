import java.util.Properties
import com.sap.conn.jco.JCoDestinationManager

module ApplicationHelper
	class SAPConnection
		def self.instance
			@_instance ||= new
		end
		
		# attr_reader :property, :conn

		def initialize
			puts("oi")
			@property = Property.new
			Settings.each do |k,v|
				@property.setProperty(k,v)
			end			

			dest = JCoDestinationManager::getDestination @property
			logger.debug("Destination attributes: " + dest.getAttributes)

		end

	end

end
