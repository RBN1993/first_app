Rails.application.routes.draw do
  # For details on the DSL available within this file, see http://guides.rubyonrails.org/routing.html
  get 'welcome/index'

  resources :users do
    get 'descarga_doc'
    get 'descarga_excel'
  end
  root 'welcome#index'

end
