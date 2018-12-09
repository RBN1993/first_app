class UsersController < ApplicationController
  skip_before_action :require_login, only: [:index, :new]

  def index
    @users = User.all
  end

  def show
    @user = User.find(params[:id])
  end

  def new

  end

  def create
    @user = User.new(user_params)
    if @user.name.present? and @user.surname.present? and @user.email.present? and @user.password.present?
      @user.save
      byebug
      redirect_to @user, success: 'Registrado correctamente'
    else
      redirect_to new_user_path, info: 'Todos los campos son obligatorios'
    end
  end

  def descarga_doc
    require 'omnidocx'
    require 'caracal'
    # for e.g. if you had to merge two documents, just pass their entire paths in an array, if you need a page break in between documents then pass the page_break flag as true
    # Omnidocx::Docx.merge_documents(['/tmp/doc1.docx', '/tmp/doc2.docx'], '/tmp/output_doc.docx', true)
    @my_data = Hash.new
    @my_data[[1, 2]] = 23
    @my_data[[5, 6]] = 42
    Caracal::Document.save 'output_doc.docx' do |docx|
      docx.page_size do
        width 15840 # sets the page width. units in twips.
        height 12240 # sets the page height. units in twips.
        orientation :landscape # sets the printer orientation. accepts :portrait and :landscape.
      end
      docx.page_margins do
        left 720 # sets the left margin. units in twips.
        right 720 # sets the right margin. units in twips.
        top 1440 # sets the top margin. units in twips.
        bottom 1440 # sets the bottom margin. units in twips.
      end
      docx.page_numbers true do
        align :center
        label ''
      end

      docx.font name: 'Courier New'

      docx.style id: 'AltFont', name: 'altFont', font: 'Palatino'
      # page 1
      docx.h1 'Page 1 Header'
      docx.hr
      docx.p
      docx.h2 'Section 1'
      docx.p 'Lorem ipsum dolor....'
      docx.p
      docx.table @my_data, border_size: 4 do
        cell_style rows[0], background: 'cccccc', bold: true
      end

      # page 2
      docx.page
      docx.h1 'Page 2 Header'
      docx.hr
      docx.p
      docx.h2 'Section 2'
      docx.p 'Lorem ipsum dolor....'
      docx.ul do
        li 'Item 1'
        li 'Item 2'
      end
      docx.p
      docx.img 'google.png', width: 500, height: 300
    end


    send_file "output_doc.docx", :filename => "output_doc.docx", :type => 'application/docx'
  end

  def descarga_excel
    require 'create_excel'

    id_json = "Hola"
    Create_excel.new_excel_images(id_json)
    send_file "/tmp/fotos_arteris_excel.xlsx", :filename => "fotos_arteris_excel.xlsx", :type => 'application/xlsx'

  end

end

private

def user_params
  params.require(:user).permit(:name,:surname,:email, :password)
end