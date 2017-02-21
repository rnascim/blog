class SapConnController < ApplicationController
	# before_action :load_from_sap

	def load
		# puts @sap_param		
		puts 'Initiate checking and loading SAP data'
		SapLoadWorker.perform_async
	end



	def init
		puts "init"
	end

end
