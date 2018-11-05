Rails.application.routes.draw do

  ### PAGINA INICIAL LOGIN ###
  root 'welcome#index'
  # root 'sessions#index'

  ### RESOURCES ###
  resources :users do
    get 'descarga_doc'
    get 'descarga_excel'
  end
  get    '/login',   to: 'sessions#new'
  post   '/login',   to: 'sessions#create'
  delete '/logout',  to: 'sessions#destroy'

end
