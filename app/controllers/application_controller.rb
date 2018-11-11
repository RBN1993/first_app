class ApplicationController < ActionController::Base
  protect_from_forgery with: :exception
  add_flash_types :danger, :info, :log_warning_on_csrf_failure, :success
  include SessionsHelper
  before_action :require_login

  private

  def require_login
    unless logged_in?
      # flash[:error] = "You must be logged in to access this section"
      # redirect_to new_user_path if current_user.blank?# halts request cycle
    end
  end
end
