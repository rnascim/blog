class CreateCompanies < ActiveRecord::Migration[5.0]
  def up
    create_table :companies, :id => false do |t|
      t.primary_key	:code, :string, :limit => 4, :null => false
      t.string :name, :limit => 40, :null => false	
      t.timestamps
    end
    add_index :companies, :code, :unique => true
  end

  def down
  	drop_table :companies
  end
end
