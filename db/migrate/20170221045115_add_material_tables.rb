class AddMaterialTables < ActiveRecord::Migration[5.0]
  
  def up
  	create_table :stloc_materials do |t|
  		t.string  "material"
  		t.string  "plant", :limit => 4
  		t.string  "stloc", :limit => 4
  		t.string  "inventory_block", :limit => 1
  		t.column  "unrestricted_stock"  , :decimal, :precision => 15, :scale => 3
  		t.column  "quality_stock", :decimal, :precision => 15, :scale => 3
  		t.column  "blocked_stock", :decimal, :precision => 15, :scale => 3
  	end

  	create_table :plant_materials do |t|
  		t.string  "material"
  		t.string  "plant", :limit => 4
  		t.string  "ncm"
  		t.string  "batch_managed", :limit => 1
  		t.string  "abc_indicator", :limit => 1
  	end

  	create_table :materials do |t|
  		t.string  "material"
  		t.string  "uom"
  		t.string  "material_group"
  		t.string  "material_type"
  	end

  	add_index :stloc_materials, [:material, :plant, :stloc]
  	add_index :plant_materials, [:material, :plant]
  	add_index :materials, [:material]
  end

  def down
  	drop_table :stloc_materials
  	drop_table :plant_materials
  	drop_table :materials
  end
end
