Rails.application.routes.draw do

  ### PAGINA INICIAL LOGIN ###
  root 'welcome#index'
  # root 'sessions#index'

  ### RESOURCES ###
  resources :users do
    get 'descarga_doc'
    get 'descarga_excel'
  end
  # get    '/login',   to: 'sessions#new' TODO: Si pasa tiempo y no se usa borrar
  post   '/login',   to: 'sessions#create' #Sirve para hacer submit en root
  delete '/logout',  to: 'sessions#destroy'

end
