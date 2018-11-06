class SessionsController < ApplicationController
  def new
  end

  def create
    user = User.find_by(email: params[:session][:email].downcase)
    if user && user.authenticate(params[:session][:password])
      log_in user
      redirect_to user, success: 'Welcome to Organizer-2018!'
    else
      # Create an error message.
      redirect_to :root, danger: 'Invalid email/password combination'

    end
  end

  def destroy
    log_out
    redirect_to root_url
  end
end
