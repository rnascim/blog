
require 'sidekiq/web'

Rails.application.routes.draw do
  resources :companies
  resources :users
  mount Sidekiq::Web, at: '/sidekiq'
  get 'sap_conn/load'

  # For details on the DSL available within this file, see http://guides.rubyonrails.org/routing.html
end
