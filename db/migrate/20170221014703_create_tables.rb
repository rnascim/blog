class CreateTables < ActiveRecord::Migration[5.0]
  def change
    create_table :tables do |t|
      t.string :name
      t.string :description

      t.timestamps
    end
    add_index :tables, :name
  end
end
