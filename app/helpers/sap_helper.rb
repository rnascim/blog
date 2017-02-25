#SapHelper is added to the applicaton and is available to any view or controller
module SapHelper

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
end #Module
