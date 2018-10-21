require 'write_xlsx'
# require 'tempfile'
require 'json'
require 'rails'
# Esta gema solo crea un excel desde cero, no es posible leer uno existente
# cuando se quiera usar imagenes dentro del excel
module Create_excel
  def self.hi
    puts "New global excel"
  end
  def self.new_excel_images(id_json)


    #Leemos del /incabidges/config.json para generar globalmente el excel
    print "HOLA"
    json = File.read("config.json")
    obj = JSON.parse(json)
    #Tener en cuenta que si no existe el uuid hay que sacar un excel con mensajes de fallo
    # id_json = "prueba_fallo"
    if(!obj[id_json])
      id_json = "no_existe"
    end

    # Libro para fotos
    fotos_book  = WriteXLSX.new("/tmp/fotos_arteris_excel.xlsx")
    fotos_sheet1 = fotos_book.add_worksheet('FOTOS 1')
    fotos_sheet2 = fotos_book.add_worksheet('FOTOS 2')

    # Formato general
    fotos_sheet1.set_zoom(110)
    fotos_sheet2.set_zoom(110)


    # byebug
    new_excel_head(fotos_book,fotos_sheet1, fotos_sheet2,obj,id_json)
    new_excel_body(fotos_book,fotos_sheet1, fotos_sheet2,obj,id_json)
    new_excel_footer(fotos_book,fotos_sheet1, fotos_sheet2,obj,id_json)

    # #foto
    # var_inv_id = InspeccionBasica.where(ib_id: params[:id]).first.inventario.inv_id
    # foto1_excel = Inventario.where(inv_id: var_inv_id).first.fotos.where(fto_sector_uuid: "515dd1e4-1353-4854-a323-4ff079309a14").first


    system('soffice --headless --convert-to xls:"MS Excel 97" /tmp/fotos_arteris_excel.xlsx --outdir /tmp/')
    fotos_book.close
    # send_file "/tmp/fotos_arteris_excel.xlsx", :filename => "fotos_arteris_excel.xlsx", :type => 'application/xlsx'
  end

  def self.new_excel_head(fotos_book, fotos_sheet1, fotos_sheet2, obj, id_json)

    # Formatos a usar
    format = fotos_book.add_format(
        :size     => 8,
        :align    => 'center',
        :valign   => 'vcenter'
    )
    format2 = fotos_book.add_format(
        :size     => 8,
        :align    => 'center',
        :valign   => 'vcenter',
        :bold     => 2,
        :border   => 5
    )

    # Foto head izq
    fotos_sheet1.insert_image('A1',"google.png",0,0,0.40,0.35)
    fotos_sheet2.insert_image('A1',"google.png",0,0,0.40,0.35)

    # Foto head derecha
    fotos_sheet1.insert_image('Q1',"google.png",350,0,0.18,0.15)
    fotos_sheet2.insert_image('Q1',"google.png",350,0,0.18,0.15)

    # Set del ancho de las correpondientes celdas
    rango = ['A:H','J:P','R:Z']
    rango2 = ['I:I','Q:Q']

    rango.each do |i|
      fotos_sheet1.set_column(i,0.50)
      fotos_sheet2.set_column(i,0.50)
    end
    rango2.each do |i|
      fotos_sheet1.set_column(i,60)
      fotos_sheet2.set_column(i,60)
    end


    # Coloco los rangos de merge dentro de un array, range2 filas con altura identica
    range = ['I1:Q1','I2:Q2','I3:Q3','I4:Q4','A6:Z6']

    range.each do |i|
      fotos_sheet1.merge_range(i,'',format)
      fotos_sheet2.merge_range(i,'',format)
    end

    fotos_sheet1.write(2, 8, obj[id_json]["title1"], format)
    fotos_sheet2.write(2, 8, obj[id_json]["title1"], format)
    fotos_sheet1.write(3,8,"Ano #{DateTime.now.year - 2007}º de Concessão", format)
    fotos_sheet2.write(3,8,"Ano #{DateTime.now.year - 2007}º de Concessão", format)

    fotos_sheet1.write(5,0, obj[id_json]["title2"],format2)
    fotos_sheet2.write(5,0, obj[id_json]["title2"],format2)

  end

  def self.new_excel_body(fotos_book,fotos_sheet1, fotos_sheet2, obj, id_json)
    # Custom body
    format = fotos_book.add_format(
        :size     => 8,
        :align    => 'center',
        :valign   => 'vcenter'
    )
    range = ['J7:P7','A8:H8','J8:P8','R8:Z8','J9:P9','A10:H10','J10:P10','R10:Z10']
    range2 = [7,9]

    range.each do |i|
      fotos_sheet1.merge_range(i,'',format)
      fotos_sheet2.merge_range(i,'',format)
    end
    range2.each do |i|
      fotos_sheet1.set_row(i,180)
      fotos_sheet2.set_row(i,180)
    end
  end

  def self.new_excel_footer(fotos_book,fotos_sheet1, fotos_sheet2, obj, id_json)
    range = [8,10]
    range.each do |i|
      fotos_sheet1.set_row(i,30)
      fotos_sheet2.set_row(i,30)
    end
  end

end