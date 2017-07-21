Rails.application.routes.draw do
	post 'home/upload'
  root to:'home#index'

  # For details on the DSL available within this file, see http://guides.rubyonrails.org/routing.html
end
