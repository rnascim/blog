class CreateValuationMaterials < ActiveRecord::Migration[5.0]
  def up
    create_table :valuation_materials do |t|
      t.string :material
      t.string :plant, :limit => 4
      t.string :valuation, :limit => 10
      t.integer :use
      t.integer :origin
      t.string :price_control
      t.decimal :moving_average_price, :precision => 13, :scale => 2
      t.decimal :standard_price, :precision => 13, :scale => 2
      t.decimal :price_unit, :precision => 18, :scale => 2
      t.decimal :stock_amount, :precision => 18, :scale => 2
      t.decimal :stock_qty, :precision => 18, :scale => 2
      t.string :status

      t.timestamps
    end

    add_index :valuation_materials, [:material, :plant, :valuation]
  end

  def down
    drop_table :valuation_materials
  end
end
