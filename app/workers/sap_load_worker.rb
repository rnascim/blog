class SapLoadWorker
  include Sidekiq::Worker
  sidekiq_options retry: false
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

	    puts "#{Time.now} - Table T001 - finished saving to local database"
  	end

  	def load_mard
  		puts "#{Time.now} - Table MARD - read begin"
  		table = sap_read_table('MARD',
  													['MATNR', 'WERKS', 'LGORT', 'SPERR', 'LABST', 'INSME', 'SPEME'],
  													["WERKS IN ('8014', '8015', 'BR11', 'BR14')"])

			puts "#{Time.now} - Table MARD - finished reading. Saving to local database"
			StlocMaterial.delete_all
	    table.each do |t|
	    	obj = StlocMaterial.new
				obj.material           = t['MATNR"]
				obj.plant              = t["WERKS']
				obj.stloc              = t['LGORT']
				obj.inventory_block    = t['SPERR']
				obj.unrestricted_stock = t['LABST']
				obj.quality_stock      = t['INSME']
				obj.blocked_stock      = t['SPEME']
				obj.save
	    end  		
	    puts "#{Time.now} - Table MARD - finished saving to local database"
  	end

  	def load_mara

  		puts "#{Time.now} - Table MARA - Assemble Material filter"
  		
  		@material_array = Array.new
  		@material_array << '('
  		materials = StlocMaterial.select('DISTINCT material')
  		materials.each do |m|
  			@material_array << "MATNR = '#{m.material}' OR"
  		end

  		if @material_array.count > 0
  			temp = @material_array[-1]
  			temp = temp[0, temp.length - 3]
  			@material_array[-1] = temp
  			@material_array << ')'
  		else
  			@material_array  = nil
  		end
  		# puts(@material_array.count)

  		puts "#{Time.now} - Table MARA - read begin"
  		table = sap_read_table('MARA',
  													['MATNR', 'MEINS', 'MATKL', 'MTART'],
  													@material_array )

  		puts "#{Time.now} - Table MARA - finished reading. Saving to local database"
			Material.delete_all
	    table.each do |t|
	    	obj = Material.new
				obj.material           = t['MATNR']
				obj.uom                = t['MEINS']
				obj.material_group     = t['MATKL']
				obj.material_type      = t['MTART']
				obj.save
	    end  		
	    puts "#{Time.now} - Table MARA - finished saving to local database"
  	end

  	def load_marc
  		puts "#{Time.now} - Table MARC - read begin"
  		table = sap_read_table('MARC',
  													['MATNR', 'WERKS', 'STEUC', 'XCHPF', 'ABCIN'],
  													["WERKS IN ('8014', '8015', 'BR11', 'BR14')"])

  		puts "#{Time.now} - Table MARC - finished reading. Saving to local database"
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
	    puts "#{Time.now} - Table MARC - finished saving to local database"
  	end

  	def load_mbew
  		puts "#{Time.now} - Table MBEW - read begin"
  		table = sap_read_table('MBEW',
  													['MATNR','BWKEY','BWTAR','MTUSE','MTORG','VPRSV','VERPR','STPRS','PEINH','SALK3','LBKUM','PSTAT'],
  													["BWKEY IN ('8014', '8015', 'BR11', 'BR14')"])

  		puts "#{Time.now} - Table MBEW - finished reading. Saving to local database"
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
	    puts "#{Time.now} - Table MBEW - finished saving to local database"
  	end

  	def load_makt
  		puts "#{Time.now} - Table MAKT - read begin (for #{@material_array.count} materials)"
  		table = sap_read_table('MAKT',
  													['MATNR','SPRAS','MAKTX'],
  													@material_array)

  		puts "#{Time.now} - Table MAKT - finished reading. Saving to local database"
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
	    puts "#{Time.now} - Table MAKT - finished saving to local database"
  	end


  	def sap_read_table(table_name, field_array, filter_array)
			java_import com.sap.conn.jco.JCoDestinationManager
			java_import com.sap.conn.jco.ext.Environment			
			java_import com.sap.conn.jco.JCoFunction
			java_import com.sap.conn.jco.JCoTable
  		time_now = Time.now			  			

			# puts("Table Name: #{table_name}" )
			# puts("Field Array: #{field_array}" )
			if !field_array == nil &&
				 !filter_array.kind_of?(Array)
				raise Exception.new("Optional parameter 'field_array' must be an Array")
			end

			# puts("Filter Array: #{filter_array}" )
			if !filter_array == nil &&
				 !filter_array.kind_of?(Array)
				raise Exception.new("Optional parameter 'filter_array' must be an Array")
			end

			dest = JCoDestinationManager.get_destination 'NFE'
			func = dest.get_repository.get_function('RFC_READ_TABLE')
			func.get_import_parameter_list.setValue('QUERY_TABLE', table_name)

			tbField = func.get_table_parameter_list.getTable('FIELDS')
			tbFilter = func.get_table_parameter_list.getTable('OPTIONS')
			tbData = func.get_table_parameter_list.getTable('DATA')

			if !field_array.nil?
				lin = 0 
				for f in field_array
					lin += 1
					# puts("#{lin} - #{f.to_s}")
					tbField.append_row
					tbField.set_row lin
					tbField.set_value 'FIELDNAME', f.to_s
				end
			end

			if !filter_array.nil?
				lin = 0
				for f in filter_array
					lin += 1
					tbFilter.append_row
					tbFilter.set_row lin
					tbFilter.set_value 'TEXT', f.to_s
				end
			end

			func.execute(dest)

			nrows = tbData.get_num_rows
			# puts("Found #{nrows} rows in the query")
			if nrows == 0
				return nil
			end

			data = Array.new
			
			nrows.times do |lin|
				tbData.set_row lin
				wa = tbData.get_string 'WA'
				pos = -1
				local_row = {}
				for f in field_array
					pos += 1
					tbField.set_row(pos)
					field_name        		 = tbField.get_string('FIELDNAME')
					start_position     		 = tbField.get_string('OFFSET').to_i
					length             		 = tbField.get_string('LENGTH').to_i
					local_field        		 = wa[start_position, length]

					# puts ("#{field_name}\t#{start_position}\t#{length}\t#{local_field}")
					local_row[field_name] = local_field
				end
				data << local_row
			end

			puts "Table #{table_name} loaded (#{data.count} rows) - Executed in #{Time.now - time_now}"
			return data

  	end
end
