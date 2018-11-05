class ApplicationController < ActionController::Base
  protect_from_forgery with: :exception
  add_flash_types :danger, :info , :log_warning_on_csrf_failure, :success
  include SessionsHelper
end
