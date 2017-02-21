		def load_from_sap
			java_import com.sap.conn.jco.JCoDestinationManager
			java_import com.sap.conn.jco.ext.Environment			
			java_import com.sap.conn.jco.JCoFunction
			
			dest = JCoDestinationManager.get_destination "NFE"

			func = dest.get_repository.get_function("");
			func.get_import_parameter_list.setValue("REQUTEXT", "Hello SAP");
			func.execute(dest)

			puts("STFC_CONNECTION finished:")
      puts(" Echo: " + func.get_export_parameter_list.get_string("ECHOTEXT"))
      puts(" Response: " +   func.get_export_parameter_list.get_string("RESPTEXT"))
			
			@sap_param = func.get_export_parameter_list


			# dest = SapJCo::Configuration.configure ({:default_destination => Rails.env})
			# tst = dest["destinations"]
			# tst.each do |k,v|
			# 	# puts(k)
			# 	if k == Rails.env
			# 		puts("\n\n>> #{k}: ------------------------(begin)")
			# 		v.each	do |k2,v2|
			# 			if k2 == "jco.client.passwd"
			# 				puts("#{k2}: \t*******")		
			# 			else
			# 				puts("#{k2}: \t#{v2}")			
			# 			end
									
			# 		end
			# 		puts("<< #{k}: ------------------------(end)\n\n")					
			# 	end
			# end

			# rfc = SapJCo::Function.new(:STFC_CONNECTION)	
			# out = rfc.execute do |params|
			#   params[:REQUTEXT] = 'Hello SAP!'
			# end
			# puts out[:ECHOTEXT] # Should print 'Hello SAP!'

			# rfc = SapJCo::Function.new(:RFC_SYSTEM_INFO)	
			# out = rfc.execute
			# puts out[:RFCSI_EXPORT][:RFCHOST2]  # Should print the host name of the SAP application server
			# @sap_param = out
		end  

		  # Create our own destination provider which converts our YAML config to a
  # java.util.Properties instance that the JCoDestinationManager can use.
  class RubyDestinationDataProvider
    include com.sap.conn.jco.ext.DestinationDataProvider
    java_import java.util.Properties
		
		def initialize
    	puts "*********** RubyDestinationDataProvider initialize: **"
    end

    def get_destination_properties(destination_name)
			java_import java.util.Properties

			puts "*********** load_from_sap"
			props = Properties.new
			props.put('jco.destination.pool_capacity'	,	"15")
			props.put('jco.client.lang'								,	"en")
			props.put('jco.client.ashost'							,	"5.189.146.49")
			props.put('jco.client.user'								,	"rogerio")
			props.put('jco.client.passwd'							,	"wolvie")
			props.put('jco.destination.peak_limit'			,	"10")
			props.put('jco.client.sysnr'								,	"00")
			props.put('jco.client.client'							,	"100")
			props.put('jco.client.trace'								,	"1")
			props.put('jco.client.abap_debug'					,	"1")

			puts "*****"
			puts props
			puts "*****"

      @properties = props
    end

    def supports_events()
      false
    end

    def set_destination_data_event_listener=(eventListener)

    end
  end
