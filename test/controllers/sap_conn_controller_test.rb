require 'test_helper'

class SapConnControllerTest < ActionDispatch::IntegrationTest
  test "should get load" do
    get sap_conn_load_url
    assert_response :success
  end

end
