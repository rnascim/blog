#source 'https://rubygems.org'

git_source(:github) do |repo_name|
  repo_name = "#{repo_name}/#{repo_name}" unless repo_name.include?("/")
  "https://github.com/#{repo_name}.git"
end


gem 'rails', '~> 5.0.1'
gem 'activerecord-jdbc-adapter', github: 'jruby/activerecord-jdbc-adapter', branch: 'rails-5'
gem 'listen'
gem 'activerecord-jdbcmysql-adapter', github: 'jruby/activerecord-jdbc-adapter'
gem 'puma', '~> 3.0'
gem 'sass-rails', '~> 5.0'
gem 'uglifier', '>= 1.3.0'
gem 'coffee-rails', '~> 4.2'
gem 'therubyrhino'
gem 'jquery-rails'
gem 'turbolinks', '~> 5'
gem 'jbuilder', '~> 2.5'
gem 'bcrypt', '~> 3.1.7'

# Use Capistrano for deployment
# gem 'capistrano-rails', group: :development


# Windows does not include zoneinfo files, so bundle the tzinfo-data gem
gem 'tzinfo-data', platforms: [:mingw, :mswin, :x64_mingw, :jruby]

# gem 'config'
gem 'sidekiq'
gem 'sinatra', require: false
gem 'slim'
gem 'devise', '~> 4.2'
gem 'simple_form', '~> 3.4'
# gem 'bootstrap', '~> 4.0.0.alpha6'

source 'https://rails-assets.org' do
  gem 'rails-assets-tether', '>= 1.3.3'
end

gem 'sprockets-rails'

#Dependency gem (for work with CentoOS)
gem 'autoprefixer-rails', github: 'ai/autoprefixer-rails'
gem 'deep_merge', github: 'danielsdeleo/deep_merge'
gem 'config', github: 'railsconfig/config'