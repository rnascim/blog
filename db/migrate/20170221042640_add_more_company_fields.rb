class AddMoreCompanyFields < ActiveRecord::Migration[5.0]
  def up
  	add_column("companies", "address", :string, :limit => 25, :after => "name")
  	add_column("companies", "country", :string, :limit => 3)
  	add_column("companies", "currency", :string, :limit => 5)
  end

  def down
  	remove_column("companies", "currency")
  	remove_column("companies", "country")
  	remove_column("companies", "address")
  end
end
