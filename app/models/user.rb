class User < ApplicationRecord
  validates :name, presence: true, length: {maximum: 4,message: "Max lo" }
  validates :email, presence: true, length: {maximum: 50 }
  validates :password,  presence: true, length: { maximum: 12 }
  has_secure_password
end
