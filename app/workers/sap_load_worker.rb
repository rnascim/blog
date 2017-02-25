class SapLoadWorker
  include Sidekiq::Worker
  include SapHelper
  sidekiq_options retry: false
  @has_assembled = false

  # sidekiq_options queue: 'high'
  

  # Rails.logger = Sidekiq::Logging.logger
  # ActiveRecord::Base.logger = Sidekiq::Logging.logger

  BATCH_NAME_SAP_LOAD = 'SAP table load'

  def perform(*args)
		puts "#{Time.now} - From Sidekiq: " + Time.now.to_s
		# newly_created = false

		batch = RunningBatch.find_by_name(BATCH_NAME_SAP_LOAD)
		if batch
			if batch.running 
				puts "#{Time.now} - Bath '#{BATCH_NAME_SAP_LOAD}' is still running"
				return
			else
				batch = RunningBatch.new
			end
		else
			batch = RunningBatch.new
			# newly_created = true
		end

		# Salva o horário de processamento do lote
		batch.name    = BATCH_NAME_SAP_LOAD
		batch.running = true
		batch.begin 	= Time.now
		batch.end   	= nil

		batch.save

		#Tabelas SAP a serem carregadas
		begin
	    load_t001
	    load_mard
	    load_marc
	    load_mbew
			load_mara
			load_makt			
		rescue Exception => e
			puts "Exception #{e.to_s}"
		end

		batch.running = false
		batch.end = Time.now
		batch.save

  end

  private
  	def load_t001
  		puts "#{Time.now} - Table T001 - read begin"
  		t001 = sap_read_table('T001',
  												 ['BUKRS', 'BUTXT', 'ORT01', 'LAND1', 'WAERS'],
  												 nil)

  		puts "#{Time.now} - Table T001 - finished reading. Saving to local database"			
			
			if !t001.nil?
				Company.delete_all
		    t001.each do |t|
		    	c = Company.new
					c.code     = t['BUKRS']
					c.name     = t['BUTXT']
					c.address  = t['ORT01']
					c.country  = t['LAND1']
					c.currency = t['WAERS']
					c.save
		    end  						
			end

	    puts "#{Time.now} - Table T001 - finished saving to local database"
  	end

  	def load_mard
  		puts "#{Time.now} - Table MARD - read begin"
  		table = sap_read_table('MARD',
  													['MATNR', 'WERKS', 'LGORT', 'SPERR', 'LABST', 'INSME', 'SPEME'],
  													["WERKS IN ('8014', '8015', 'BR11', 'BR14')"])

			puts "#{Time.now} - Table MARD - finished reading. Saving to local database"
			if !table.nil?
				StlocMaterial.delete_all
		    table.each do |t|
		    	obj = StlocMaterial.new
					obj.material           = t['MATNR']
					obj.plant              = t['WERKS']
					obj.stloc              = t['LGORT']
					obj.inventory_block    = t['SPERR']
					obj.unrestricted_stock = t['LABST']
					obj.quality_stock      = t['INSME']
					obj.blocked_stock      = t['SPEME']
					obj.save
		    end  		
			end
	    puts "#{Time.now} - Table MARD - finished saving to local database"
  	end

  	def assemble_material_array
  		if @has_assembled
  			puts "#{Time.now} - Assemble Material filter (return with content in memory)"
  			return	
  		end

  		@has_assembled = true

			puts "#{Time.now} - Assemble Material filter (assembling)"
  		@material_array = Array.new
  		@material_array << '('
  		materials = StlocMaterial.select('DISTINCT material')
  		if !materials.nil?
	  		materials.each do |m|
	  			@material_array << "MATNR = '#{m.material}' OR"
	  		end  			
  		end

			if @material_array.count > 1
  			temp = @material_array[-1]
  			temp = temp[0, temp.length - 3]
  			@material_array[-1] = temp
  			@material_array << ')'
  		else
  			@material_array  = nil
  		end

  	end

  	def load_mara
  		
  		@material_array = assemble_material_array

  		puts "#{Time.now} - Table MARA - read begin"
  		table = sap_read_table('MARA',
  													['MATNR', 'MEINS', 'MATKL', 'MTART'],
  													@material_array )

  		puts "#{Time.now} - Table MARA - finished reading. Saving to local database"
			if !table.nil?
				Material.delete_all
		    table.each do |t|
		    	obj = Material.new
					obj.material           = t['MATNR']
					obj.uom                = t['MEINS']
					obj.material_group     = t['MATKL']
					obj.material_type      = t['MTART']
					obj.save
		    end  						
			end
	    puts "#{Time.now} - Table MARA - finished saving to local database"
  	end

  	def load_marc
  		puts "#{Time.now} - Table MARC - read begin"
  		table = sap_read_table('MARC',
  													['MATNR', 'WERKS', 'STEUC', 'XCHPF', 'ABCIN'],
  													["WERKS IN ('8014', '8015', 'BR11', 'BR14')"])

  		puts "#{Time.now} - Table MARC - finished reading. Saving to local database"
  		if !table.nil?
				PlantMaterial.delete_all
		    table.each do |t|
		    	obj = PlantMaterial.new
					obj.material           = t['MATNR']
					obj.plant              = t['WERKS']
					obj.ncm    						 = t['STEUC']
					obj.batch_managed      = t['XCHPF']
					obj.abc_indicator      = t['ABCIN']
					obj.save
		    end  		
  		end
	    puts "#{Time.now} - Table MARC - finished saving to local database"
  	end

  	def load_mbew
  		puts "#{Time.now} - Table MBEW - read begin"
  		table = sap_read_table('MBEW',
  													['MATNR','BWKEY','BWTAR','MTUSE','MTORG','VPRSV','VERPR','STPRS','PEINH','SALK3','LBKUM','PSTAT'],
  													["BWKEY IN ('8014', '8015', 'BR11', 'BR14')"])

  		puts "#{Time.now} - Table MBEW - finished reading. Saving to local database"
			if !table.nil?
				ValuationMaterial.delete_all
		    table.each do |t|
		    	obj = ValuationMaterial.new
					obj.material             = t['MATNR']
					obj.plant                = t['BWKEY']
					obj.valuation					   = t['BWTAR']
					obj.use                  = t['MTUSE']
					obj.origin               = t['MTORG']
					obj.price_control		     = t['VPRSV']
					obj.moving_average_price = t['VERPR']
					obj.standard_price			 = t['STPRS']
					obj.price_unit				   = t['PEINH']
					obj.stock_amount				 = t['SALK3']
					obj.stock_qty    				 = t['LBKUM']
					obj.status							 = t['PSTAT']
					obj.save
		    end  						
			end
	    puts "#{Time.now} - Table MBEW - finished saving to local database"
  	end

  	def load_makt
  		
  		assemble_material_array

  		puts "#{Time.now} - Table MAKT - read begin (for #{@material_array} materials)"
  		table = sap_read_table('MAKT',
  													['MATNR','SPRAS','MAKTX'],
  													@material_array)

  		puts "#{Time.now} - Table MAKT - finished reading. Saving to local database"
  		if !table.nil?
				MaterialName.delete_all
		    table.each do |t|
		    	obj = MaterialName.new
					obj.material             = t['MATNR']
					case t['SPRAS']
					when 'P'
						obj.language         = 'PT'
					when 'D'
						obj.language         = 'DE'
					when 'E'
						obj.language         = 'EN'
					when 'S'
						obj.language         = 'ES'
					when 'F'
						obj.language         = 'FR'
					when '3'
						obj.language         = 'KO'
					when 'B'
						obj.language         = 'HE'
					when 'T'
						obj.language         = 'TR'
					when '1'
						obj.language         = 'ZH'
					when 'I'
						obj.language         = 'IT'
					when 'L'
						obj.language         = 'PL'
					else
						obj.language         = t['SPRAS']
					end
					obj.name    					   = t['MAKTX']
					obj.save
		    end  		
		  end
	    puts "#{Time.now} - Table MAKT - finished saving to local database"
  	end

end
