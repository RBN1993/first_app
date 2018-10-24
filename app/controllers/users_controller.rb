class UsersController < ApplicationController

  def show
    @user = User.find(params[:id])
  end

  def new
  end

  def create
    @user = User.new(user_params)

    @user.save
    redirect_to @user
  end

  def descarga_doc
    require 'omnidocx'
    require 'caracal'
    # for e.g. if you had to merge two documents, just pass their entire paths in an array, if you need a page break in between documents then pass the page_break flag as true
    # Omnidocx::Docx.merge_documents(['/tmp/doc1.docx', '/tmp/doc2.docx'], '/tmp/output_doc.docx', true)
    docx = Caracal::Document.new('/tmp/output_doc.docx')
    docx.page_size do
      width       15840       # sets the page width. units in twips.
      height      12240       # sets the page height. units in twips.
      orientation :landscape  # sets the printer orientation. accepts :portrait and :landscape.
    end
    send_file "/tmp/output_doc.docx", :filename => "output_doc.docx", :type => 'application/docx'
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
  params.require(:user).permit(:name, :surname)
end