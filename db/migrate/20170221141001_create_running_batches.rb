class CreateRunningBatches < ActiveRecord::Migration[5.0]
  def up
    create_table :running_batches do |t|
    	t.string	"name"
    	t.boolean "running"
    	t.datetime "begin"
    	t.datetime "end"
      t.timestamps
    end

    add_index :running_batches, :name
  end

  def down
  	drop_table :running_batches
  end
end
