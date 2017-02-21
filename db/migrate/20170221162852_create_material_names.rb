class CreateMaterialNames < ActiveRecord::Migration[5.0]
  def up
    create_table :material_names do |t|
      t.string :material
      t.string :language
      t.string :name

      t.timestamps
    end
    add_index :material_names, [:language, :material]
  end

  def down
  	drop_table :material_names
  end
end
