class InspeccionesController < ApplicationController
  include InspeccionesHelper
  include IndicesDesglosados if CUSTOMER.config.indices_desglosados
  include IndicesContencion if CUSTOMER.config.indices_desglosados
  require 'zip'
  before_filter :require_user
  before_filter :load_inventario

  def index
    load_inventario
    if SETTINGS.inspecciones_principales_enabled?
    @last_ip = @inventario.inspecciones_principales.where("ip_validado = TRUE").order("ip_fechainsp DESC").first
    end
    if @last_ip.present?
      datos = @last_ip.datos_elementos
      subcomponentes = datos.map { |x| [x[0].tipologia.translate, x[0].tipologia.uuid] }
      elementos = datos.map { |x| [x[1], x[1]] }
      unidad = datos.map { |x|  x[2] }
      @datos_filtro = []
      @datos_filtro[0] = @last_ip.elementos.map{|x| n = MASTER_DATA[x.com_deterioro_uuid]; [n.try(:translate), n.try(:uuid)] }.uniq
      @datos_filtro[1] = @last_ip.elementos.map{|x| n = MASTER_DATA[x.com_gravedad_uuid]; [n.try(:translate), n.try(:uuid)] }.uniq
      @datos_filtro[2] = @last_ip.elementos.map{|x| n = MASTER_DATA[x.com_causa_uuid]; [n.try(:translate), n.try(:uuid)] }.uniq
      @datos_filtro[3] = @last_ip.elementos.map{|x| n = MASTER_DATA[x.com_extension_uuid]; [n.try(:translate), n.try(:uuid)] }.uniq
      @datos_filtro[4] = subcomponentes.uniq
      @datos_filtro[5] = elementos.uniq
      @datos_filtro[6] = unidad.flatten(1).uniq
    end
    render_default
  end

  def show
    fetch_inspeccion
    @horizontal = params[:horizontal] || SETTINGS.mop?

    respond_to do |format|
      format.html {
        @visto_bueno = @inspeccion.validado || params[:validate].present?
        @inspeccion.calculate_calification if @inspeccion.respond_to?(:calculate_calification)
        @errors = @inspeccion.errors
        render_default
      }
      format.print {
        @inventario = @inspeccion.inventario
        filtered_elements = inspection_filter_report(@inspeccion)
        render report_template, :layout => 'print.html.haml', :locals => {:horizontal => @horizontal, :title => "#{@inspeccion.prefix}-#{@inventario.inv_codigopuente}", :elements => filtered_elements}
      }
      format.zip {
        codigo = @inspeccion.gen_codigo
        # byebug
        stringio = Zip::OutputStream.write_buffer do |zio|
          #prefix = ''
          prefix = codigo + '/'
          zio.put_next_entry("#{prefix}#{codigo}.xml")
          zio.write @inspeccion.to_xml(:filenames => true)
          @inspeccion.fotografias.each do |f|
            zio.put_next_entry("#{prefix}foto/"+f.data_file_name)
            zio.write((File.exists?(f.data.path) && Paperclip.io_adapters.for(f.data) || File.new("#{Rails.root}/public/images/missing_file.jpg", "r")).read)
          end
          @inspeccion.grabaciones.each do |f|
            zio.put_next_entry("#{prefix}audio/"+f.data_file_name)
            zio.write((File.exists?(f.data.path) && Paperclip.io_adapters.for(f.data) || File.new("#{Rails.root}/public/images/missing_file.jpg", "r")).read)
          end
          @inspeccion.elementos.each do |el|
            el.fotografias.each do |f|
              zio.put_next_entry("#{prefix}foto/"+f.data_file_name)
              zio.write((File.exists?(f.data.path) && Paperclip.io_adapters.for(f.data) || File.new("#{Rails.root}/public/images/missing_file.jpg", "r")).read)
            end
            el.grabaciones.each do |f|
              zio.put_next_entry("#{prefix}audio/"+f.data_file_name)
              zio.write((File.exists?(f.data.path) && Paperclip.io_adapters.for(f.data) || File.new("#{Rails.root}/public/images/missing_file.jpg", "r")).read)
            end
          end
          #zio.write "\xFF\xFE" + Iconv.conv("utf-16le", "utf-8", @inspeccion.to_xml)
        end
        stringio.rewind
        data = stringio.sysread

        filename = codigo + ".zip"
        headers["Content-Disposition"] = "attachment; filename=\"#{filename}\""
        send_data data, :type => 'text/zip', :disposition => 'attachment', :filename => filename

      }
      format.pdf do
        @inventario = @inspeccion.inventario
        filtered_elements = inspection_filter_report(@inspeccion)
        render file: report_template,
               pdf: nombre_informe(@inspeccion, :inspeccion),
               layout: 'pdf.html',
               locals: {
                   horizontal: @horizontal,
                   title: "#{@inspeccion.prefix}-#{@inventario.inv_codigopuente}",
                   elements: filtered_elements
               }
      end
      format.json do
        render :show
      end
    end
  end

  def edit
    show
  end

  def new
    if params[:media] == 'zip'
      @inspeccion = build_new
      @import_zip = true
      render_default
    else
      @inspeccion = build_for_structure
      if @inspeccion.save
        redirect_to inspeccion_path
      else
        render_default
      end
    end
  end

  def update
    fetch_inspeccion
    inspeccion_params
    if @inspeccion.update_attributes(inspeccion_params)
      @inspeccion.calculate_calification
      @inspeccion.save!
      if CUSTOMER.config.indices_desglosados && @inspeccion.class.name == 'InspeccionPrincipal'
        @inspeccion.calcula_indices(@inspeccion)
        flash[:notice] = '' if flash[:notice].nil?
        if SETTINGS.a4?
          flash[:notice] << ' --A TOMAR EN CUENTA SOBRE INDICE CONTENCIÓN-- '
        else
          flash[:notice] << ' --A TOMAR EN CUENTA SOBRE INDICE CONTENCIÓN-- '
        end
        flash[:notice] << @@contencion_errors.join("\n")
      end
      flash[:error] = ""
      calculate_next_inspections rescue flash[:error] << "Error al calcular el calendario de inspecciones      \n"
      calcular_apme if (SETTINGS.arteris? || SETTINGS.efe? || SETTINGS.chile? || SETTINGS.mcp? || SETTINGS.a4?) && @inspeccion.is_a?(InspeccionPrincipal)
      respond_to do |format|
        format.html {
          if @inspeccion.calification_errors.present?
            flash[:error] << "Calificación: #{@inspeccion.calification_errors.join("\n")}"
            flash[:error] = nil if flash[:error].blank?
          else
            flash[:notice] = I18n.t("inventario.update.success")
          end
          redirect_to inspeccion_path
        }
        format.json {
          if @inspeccion.calification_errors.present?
            render json: {id: @inspeccion.id, @inspeccion.class.name.underscore => serialize_inspeccion(@inspeccion), :error => "Calificación: #{@inspeccion.calification_errors.join("\n")}"}
          else
            render json: {id: @inspeccion.id, @inspeccion.class.name.underscore => serialize_inspeccion(@inspeccion), :notice => MD_TEXTO_GUARDADO.to_s}
          end
        }
      end
    else
      respond_to do |format|
        format.html {
          render_default
        }
        format.json {
          puts "*"*80
          puts @inspeccion.errors.inspect
          puts "*"*80
          render :json => {errors: @inspeccion.errors.as_json}, status: :unprocessable_entity
        }
      end
    end
  end

  def validate
    fetch_inspeccion
    @inspeccion.validado = true
    @inspeccion.validador = current_user
    @inspeccion.fecha_validacion = Time.now

    if @inspeccion.save
      unless @inspeccion.is_a? InspeccionEspecial
        Mailer.inspection_report(@inspeccion) if SETTINGS.mail_notifications_enabled? && @inspeccion.need_report?
      end
      flash[:notice] = I18n.t("inventario.create.validated")
    end
    redirect_to :back
  end

  def unvalidate
    fetch_inspeccion
    @inspeccion.validado = false
    @inspeccion.validador = current_user
    @inspeccion.fecha_validacion = Time.now

    if @inspeccion.save
      flash[:notice] = I18n.t("inventario.create.unvalidated")
    end
    redirect_to inspeccion_path
  end

  def create
    @inventario = Inventario.find(inspeccion_params[:inv_id])
    if params[:file]
      begin
        @inspeccion = build_from_zip
        if @inspeccion.inv_id != @inventario.id
          flash[:error] = "La inspección para la estructura #{@inspeccion.inventario.inv_codigopuente} no se corresponde con el inventario seleccionado"
          redirect_to :back
          return
        end
      rescue => e
        flash[:error] = e.message
        redirect_to :back
        return
      end
    else
      @inspeccion = build_new(inspeccion_params)
    end

    @inspeccion.calculate_calification

    if @inspeccion.save
      calculate_next_inspections

      #if @inspeccion.calification_errors.present?
      #  flash[:error] = "Calificación: #{@inspeccion.calification_errors.join("\n")}"
      #else
      flash[:notice] = I18n.t("inventario.create.success")
      #end
      redirect_to inspeccion_path
    else
      if params[:file]
        @import_zip = true
      end
      render_default
    end
  end

  def destroy
    fetch_inspeccion
    inv_id = @inspeccion.inventario.inv_id
    @inventario = Inventario.find(inv_id)
    @inspeccion.destroy
    @inventario.calculate_next_inspections if (SETTINGS.arteris? || SETTINGS.efe? || SETTINGS.chile?)
    respond_to do |format|
      format.html {
        flash[:notice] = I18n.t("inventario.destroy.success")
        redirect_to inspecciones_path(:inventario => @inventario)
      }
      format.json {
        render :json => {notice: "Inspección eliminada", redirect: inspecciones_path(:inventario => @inventario)}
      }
    end
  end

  def load_inventario
    @inventario ||= Inventario.find(params[:inventario]) if params[:inventario].present?
    authorize! action_name.to_sym, @inspeccion if @inspeccion.present?
  end

  def fetch_data_for_layout
    @inventario = @inspeccion.inventario if @inspeccion
    @section = MD_SECCION_INSPECCIONES
    @tunel = TunelInspeccion.where(:inv_id => @inventario).select { |i| can?(:read, i) } if SETTINGS.inspecciones_tuneles_enabled?
    @basic = InspeccionBasica.where(:inv_id => @inventario).select { |i| can?(:read, i) } if SETTINGS.inspecciones_basicas_enabled?
    @main = InspeccionPrincipal.where(:inv_id => @inventario).select { |i| can?(:read, i) } if SETTINGS.inspecciones_principales_enabled?
    @special = InspeccionEspecial.where(:inv_id => @inventario).select { |i| can?(:read, i) } if CUSTOMER.config.inspeccion_especial
  end

  def render_default
    fetch_data_for_layout
    render 'layouts/_inspecciones'
  end

  def calculate_next_inspections
    if !@inspeccion.try(:ib_evento_asociado).nil? && @inspeccion.try(:ib_evento_asociado) != 0 && (SETTINGS.efe? || SETTINGS.chile? || SETTINGS.arteris? || SETTINGS.mcp? || SETTINGS.a4?)
      @inspeccion.calculate_next_inspections
    else
      if SETTINGS.calendario_inspecciones_enabled?
        if (!@inspeccion.inventario.calculate_next_inspections rescue false)
          flash[:error] = "Error al calcular próximas inspecciones"
        end
      end
    end
  end

  def add_informe_excel
    require 'spreadsheet'
    require 'geo/coord'
    require 'rake'
    require 'create_excel'

    @ib= InspeccionBasica.where(ib_id: params[:id]).first
    var_ib_id = @ib.ib_id
    var_inv_id = @ib.inventario.inv_id
    @inv= Inventario.where(inv_id: var_inv_id).first

    var_codigo_antt = @inv.inv_codigopuente_antt
    var_nombre_puente = @inv.inv_nombrepuente
    var_carretera = @inv.inv_fun_carretera
    var_inv_pk_km = @inv.inv_pk
    var_inv_pk_mts = @inv.inv_pkmts
    var_ib_fecha = @ib.ib_fechainsp.to_date
    var_estabilidad = @ib.ib_estabilidad_uuid
    var_conservacion = @ib.ib_conservacion_uuid
    var_vibra_tablero = @ib.ib_vibracion_tablero_uuid
    var_ins_previas = @ib.ib_inspecciones_previas
    var_ins_necesaria = @ib.ib_insp_esp_necesaria
    var_aviso_urgente = @ib.ib_avisourgente
    var_hist_interven = @ib.ib_historico_intervencion
    var_observacion = @ib.ib_observaciones
    #var_gravedad_det = @ib.elementos.map{|x| x.com_gravedad_uuid}
    
    var_nota_tecnica = SubinformeIb.where(ib_id: var_ib_id).order(:sub_id).map{|e| e.elementos.map{|x| MASTER_DATA[x.com_gravedad_uuid].value}.max}
    var_detalle_uno = SubinformeIb.where(ib_id: var_ib_id).map{|s| s.elementos.map{|x| [x.detalle_uuid]}}
    var_detalle_dos = SubinformeIb.where(ib_id: var_ib_id).map{|s| s.elementos.map{|x| [x.detalle2_uuid]}}
    var_detalle_tres = SubinformeIb.where(ib_id: var_ib_id).map{|s| s.elementos.map{|x| [x.detalle3_uuid]}}
    var_coordenadas1 = @inv.inv_geometria.split(":")[1].split("|").to_a[0].split(",").to_a[0]
    var_coordenadas2 = @inv.inv_geometria.split(":")[1].split("|").to_a[0].split(",").to_a[1]
    var_comentario_localizacion = SubinformeIb.where(ib_id: var_ib_id).map{|s| s.elementos.map{|x| [x.com_localizacion]}}
    var_comentario_medicion = SubinformeIb.where(ib_id: var_ib_id).map{|s| s.elementos.map{|x| [x.com_medicion]}}
    var_comentario_observaciones = SubinformeIb.where(ib_id: var_ib_id).map{|s| s.elementos.map{|x| [x.com_observaciones]}}
    var_ud = SubinformeIb.where(ib_id: var_ib_id).map{|s| s.elementos.map{|x| [x.com_unidad_medicion]}}

    #SACAR COORDENADA
    var_coord_final = Geo::Coord.new(lat:var_coordenadas1, lng:var_coordenadas2)

    #SACAR CLAVE ESTADO
    sacar_estado = @inv.inv_regiones_uuid
    if sacar_estado == "0d7841cb-f2cc-446e-a3f8-c3b50a36c73c"
      var_iniciales_estado = "AC"
    elsif sacar_estado == "143951e2-e1e8-4d68-8e8b-3a42016c6270"
      var_iniciales_estado = "AL"
    elsif sacar_estado == "bde2fe0a-119b-4076-abe3-ab160740e9a9"
      var_iniciales_estado = "AP"
    elsif sacar_estado == "74c97b00-2844-480e-8a6a-25fed4754f73"
      var_iniciales_estado = "AM"
    elsif sacar_estado == "37ebfd63-4d4c-409d-b617-8e8905fca24e"
      var_iniciales_estado = "BA"
    elsif sacar_estado == "95340a12-c8e9-4b3a-bf6d-75cfeb55e033"
      var_iniciales_estado = "CE"
    elsif sacar_estado == "8cb53147-2634-4ba5-90c1-b004b3dbf02b"
      var_iniciales_estado = "ES"
    elsif sacar_estado == "80c2f482-2c5e-47c2-9aeb-de17adb2df3f"
      var_iniciales_estado = "GO"
    elsif sacar_estado == "87895d55-fb15-43a6-bba7-cd31cdd4eb7f"
      var_iniciales_estado = "MA"
    elsif sacar_estado == "0ce8223c-2153-4303-b1a7-7ecd5a80c0f6"
      var_iniciales_estado = "MT"
    elsif sacar_estado == "9a49c6cd-3c8a-4e02-b58d-f3f4e09f14b6"
      var_iniciales_estado = "MS"
    elsif sacar_estado == "c2826d5d-f011-484b-951e-7805c9ee8fe0"
      var_iniciales_estado = "MG"
    elsif sacar_estado == "c9f09ed8-2ff0-4db4-846e-6f86182e7139"
      var_iniciales_estado = "PA"
    elsif sacar_estado == "06a18943-f91e-4227-8c8f-0962c40116f7"
      var_iniciales_estado = "PB"
    elsif sacar_estado == "9404b496-46d8-40ed-bdb8-94a8af1bd6a5"
      var_iniciales_estado = "PR"
    elsif sacar_estado == "87ec1ce1-fa16-4e3f-b6f1-8287a8aca627"
      var_iniciales_estado = "PE"
    elsif sacar_estado == "ed7b9322-c00b-40e5-a246-f4383b4b7551"
      var_iniciales_estado = "PI"
    elsif sacar_estado == "ade8745f-830e-426d-b3c2-530cd7df33b6"
      var_iniciales_estado = "RJ"
    elsif sacar_estado == "cd7a9f73-3da2-493b-8b30-d41043bed4ba"
      var_iniciales_estado = "RN"
    elsif sacar_estado == "fc15fe60-f4fc-4ca2-9f57-ec4c1a79870e"
      var_iniciales_estado = "RS"
    elsif sacar_estado == "37563d90-a2c1-460d-92f8-8c35fa8a68b0"
      var_iniciales_estado = "RO"
    elsif sacar_estado == "38466dfa-76a0-497b-ab9d-5bd7a4d2c9b8"
      var_iniciales_estado = "RR"
    elsif sacar_estado == "d7fa3226-f1e4-45b9-8ed4-23cd7a254a51"
      var_iniciales_estado = "SC"
    elsif sacar_estado == "2500c6bf-9288-4889-8ae0-b9a62198ae39"
      var_iniciales_estado = "SP"
    elsif sacar_estado == "92bf319b-e21a-4c6d-8172-3b641862507d"
      var_iniciales_estado = "TO"
    elsif sacar_estado == "38f6b195-584a-4a2c-9e44-44f7b078c01d"
      var_iniciales_estado = "DF"
    end

    #libro para escibir
    if @inv.inv_concesionaria_uuid == "aaa189b7-133b-413a-9dae-b08918f62485"
      book = "#{Rails.root}/custom/#{CUSTOMER.name}/tables/ficha_dias.xls"
    elsif @inv.inv_concesionaria_uuid == "1f12eb97-b4a1-4e31-964f-c7f5eb195c34"
      book = "#{Rails.root}/custom/#{CUSTOMER.name}/tables/ficha_fluminense.xls"
    elsif @inv.inv_concesionaria_uuid == "bc8d8d68-303c-4efc-86d7-14d5398c90aa"
      book = "#{Rails.root}/custom/#{CUSTOMER.name}/tables/ficha_litoral.xls"
    elsif @inv.inv_concesionaria_uuid == "250210de-8549-4aee-94a7-6e3ae22bcc47"
      book = "#{Rails.root}/custom/#{CUSTOMER.name}/tables/ficha_planato.xls"
    elsif @inv.inv_concesionaria_uuid == "8ca3bb15-57b3-4b10-ae57-58d7b8f3cee5"
      book = "#{Rails.root}/custom/#{CUSTOMER.name}/tables/ficha_regis.xls"
    end
    id_json = @inv.inv_concesionaria_uuid
    open_book = Spreadsheet.open(book)
    hoja_trabajo = open_book.worksheet(0)
    hoja_trabajo1 = open_book.worksheet(1)
    hoja_trabajo2 = open_book.worksheet(2)
      # byebug
      hoja_trabajo.rows[3][3] = "Ano #{DateTime.now.year - 2007}º de Concessão"
      hoja_trabajo1.rows[3][8] = "Ano #{DateTime.now.year - 2007}º de Concessão"
      hoja_trabajo2.rows[3][8] = "Ano #{DateTime.now.year - 2007}º de Concessão"
      hoja_trabajo.rows[5][2] = var_codigo_antt
      hoja_trabajo.rows[5][7] = var_nombre_puente
      hoja_trabajo.rows[5][24] = "#{var_carretera}" "/" "#{var_iniciales_estado}"
      # byebug
      length=var_inv_pk_km.to_s.length
      var_inv_pk_km_str = '0' * (3-length)  + var_inv_pk_km.to_s  unless length>3 
      
      length=var_inv_pk_mts.to_s.length
      var_inv_pk_mts_str = '0' * (3-length)  + var_inv_pk_mts.to_s  unless length>3 

      hoja_trabajo.rows[5][28] = "#{var_inv_pk_km_str}" "+" "#{var_inv_pk_mts_str}"
      hoja_trabajo.rows[7][1] = var_ib_fecha.strftime("%d/%m/%Y")
      #hoja_trabajo.rows[5][16] = var_coord_final.to_s
      hoja_trabajo.rows[6][16] = var_coord_final.to_s.split(" ")[0]
      hoja_trabajo.rows[7][16] = var_coord_final.to_s.split(" ")[1]
      #Condiciones de estabilidad
      if var_estabilidad == "4cbca4e7-2f26-4763-96b5-8d5ea00adb99"
        hoja_trabajo.rows[11][4] = "X"
      elsif var_estabilidad == "31f8c352-72d1-4cb0-a977-08f6ef7eff81"
        hoja_trabajo.rows[11][9] = "X"
      elsif var_estabilidad == "0855ac22-43d6-444a-8325-cd7b59647491"
        hoja_trabajo.rows[11][13] = "X"
      end
      #Condiciones de conservacion
      if var_conservacion == "2d3b0342-a556-4a0f-9510-d873b1a6a7ae"
        hoja_trabajo.rows[12][4] = "X"
      elsif var_conservacion == "86a21837-e38f-43fe-bcf6-d36afa10a695"
        hoja_trabajo.rows[12][9] = "X"
      elsif var_conservacion == "3d1bedf1-953b-48b8-9b3c-7f0c6fcd27cb"
        hoja_trabajo.rows[12][13] = "X"
      elsif var_conservacion == "c48b1d3d-5512-48af-aed4-813099efc026"
        hoja_trabajo.rows[12][18] == "X"
      end
      #Nivel de vibracion tablero
      if var_vibra_tablero == "a3bda573-d52a-4f18-9358-7c9112b11395"
        hoja_trabajo.rows[13][4] = "X"
      elsif var_vibra_tablero == "826683ba-a167-469a-952c-1fe0447258ba"
        hoja_trabajo.rows[13][9] = "X"
      elsif var_vibra_tablero == "f92b532a-1045-4545-a03b-67c4c61ad1ca"
        hoja_trabajo.rows[13][13] = "X"
      end
      #Inspeccion previa
      if var_ins_previas == true
        hoja_trabajo.rows[15][4] = "X"
      elsif var_ins_previas == false
        hoja_trabajo.rows[15][9] = "X"
      end
      #Inspeccion necesaria
      if var_ins_necesaria == true
        hoja_trabajo.rows[14][18] = "X"
      elsif var_ins_necesaria == false
        hoja_trabajo.rows[14][22] = "X"
      end
      #Aviso urgente
      if var_aviso_urgente == true
        hoja_trabajo.rows[14][29] = "X"
      elsif var_aviso_urgente == false
        hoja_trabajo.rows[14][34] = "X"
      end
      #Comentario
      hoja_trabajo.rows[17][12] = var_hist_interven
      #Observacion datos generales
      hoja_trabajo.rows[20][0] = var_observacion

      #LAJE
        #Nota tecnica
        if var_nota_tecnica[0] == "4"
          hoja_trabajo.rows[40][17] = 1
        elsif var_nota_tecnica[0] == "3"
          hoja_trabajo.rows[40][17] = 2
        elsif var_nota_tecnica[0] == "2"
          hoja_trabajo.rows[40][17] = 3
        elsif var_nota_tecnica[0] == "1"
          hoja_trabajo.rows[40][17] = 4
        elsif var_nota_tecnica[0] == "0"
          hoja_trabajo.rows[40][17] = 5
        end
        #Deterioro1
          unless var_detalle_uno[0][0] == [nil]
            hoja_trabajo.rows[41][4] = "X"
          end
          unless var_detalle_dos[0][0] == [nil]
          hoja_trabajo.rows[41][13] = "X"
          end
          unless var_comentario_localizacion[0][0] == nil
            hoja_trabajo.rows[41][20] = var_comentario_localizacion[0][0].join
          end
          unless var_comentario_medicion[0][0] == nil
            hoja_trabajo.rows[41][36] = "#{var_comentario_medicion[0][0].first.to_f} #{var_ud[0][0].join}"
          end
        #Deterioro2
          unless var_detalle_uno[0][1] == [nil]
            hoja_trabajo.rows[42][4] = "X"
          end
          unless var_detalle_dos[0][1] == [nil]
          hoja_trabajo.rows[42][13] = "X"
          end
          unless var_comentario_localizacion[0][1] == nil
            hoja_trabajo.rows[42][20] = var_comentario_localizacion[0][1].join
          end
          unless var_comentario_medicion[0][1] == nil
            hoja_trabajo.rows[42][36] = "#{var_comentario_medicion[0][1].first.to_f} #{var_ud[0][1].join}"
          end
        #Deterioro3
          unless var_detalle_uno[0][2] == [nil]
            hoja_trabajo.rows[43][4] = "X"
          end
          unless var_detalle_dos[0][2] == [nil]
          hoja_trabajo.rows[43][13] = "X"
          end
          unless var_comentario_localizacion[0][2] == nil
            hoja_trabajo.rows[43][20] = var_comentario_localizacion[0][2].join
          end
          unless var_comentario_medicion[0][2] == nil
            hoja_trabajo.rows[43][36] = "#{var_comentario_medicion[0][2].first.to_f} #{var_ud[0][2].join}"
          end
        #Deterioro4
          unless var_detalle_uno[0][3] == [nil]
            hoja_trabajo.rows[44][4] = "X"
          end
          unless var_detalle_dos[0][3] == [nil]
          hoja_trabajo.rows[44][13] = "X"
          end
          unless var_comentario_localizacion[0][3] == nil
            hoja_trabajo.rows[44][20] = var_comentario_localizacion[0][3].join
          end
          unless var_comentario_medicion[0][3] == nil
            hoja_trabajo.rows[44][36] = "#{var_comentario_medicion[0][3].first.to_f} #{var_ud[0][3].join}"
          end
        #Deterioro5
          unless var_detalle_uno[0][4] == [nil]
            hoja_trabajo.rows[45][4] = "X"
          end
          unless var_detalle_dos[0][4] == [nil]
          hoja_trabajo.rows[45][13] = "X"
          end
          unless var_comentario_localizacion[0][4] == nil
            hoja_trabajo.rows[45][20] = var_comentario_localizacion[0][4].join
          end
          unless var_comentario_medicion[0][4] == nil
            hoja_trabajo.rows[45][36] = "#{var_comentario_medicion[0][4].first.to_f} #{var_ud[0][4].join}"
          end
        #Deterioro6
          unless var_detalle_uno[0][5] == [nil]
            hoja_trabajo.rows[46][4] = "X"
          end
          unless var_comentario_localizacion[0][5] == nil
            hoja_trabajo.rows[46][20] = var_comentario_localizacion[0][5].join
          end
          unless var_comentario_medicion[0][5] == nil
            hoja_trabajo.rows[46][36] = "#{var_comentario_medicion[0][5].first.to_f} #{var_ud[0][5].join}"
          end
        #Deterioro7
          unless var_detalle_uno[0][6] == [nil]
            hoja_trabajo.rows[47][4] = "X"
          end
          unless var_comentario_localizacion[0][6] == nil
            hoja_trabajo.rows[47][20] = var_comentario_localizacion[0][6].join
          end
          unless var_comentario_medicion[0][6] == nil
            hoja_trabajo.rows[47][36] = "#{var_comentario_medicion[0][6].first.to_f} #{var_ud[0][6].join}"
          end
      #VIGAMENTO PRINCIPAL
        #Nota tecnica
        if var_nota_tecnica[1] == "4"
          hoja_trabajo.rows[49][17] = 1
        elsif var_nota_tecnica[1] == "3"
          hoja_trabajo.rows[49][17] = 2
        elsif var_nota_tecnica[1] == "2"
          hoja_trabajo.rows[49][17] = 3
        elsif var_nota_tecnica[1] == "1"
          hoja_trabajo.rows[49][17] = 4
        elsif var_nota_tecnica[1] == "0"
          hoja_trabajo.rows[49][17] = 5
        end
        #Deterioro1
          unless var_detalle_uno[1][0] == [nil]
            hoja_trabajo.rows[50][4] = "X"
          end
          unless var_detalle_dos[1][0] == [nil]
          hoja_trabajo.rows[50][13] = "X"
          end
          unless var_comentario_localizacion[1][0] == nil
            hoja_trabajo.rows[50][20] = var_comentario_localizacion[1][0].join
          end
          unless var_comentario_medicion[1][0] == nil
            hoja_trabajo.rows[50][36] = "#{var_comentario_medicion[1][0].first.to_f} #{var_ud[1][0].join}"
          end
        #Deterioro2
          unless var_detalle_uno[1][1] == [nil]
            hoja_trabajo.rows[51][4] = "X"
          end
          unless var_detalle_dos[1][1] == [nil]
          hoja_trabajo.rows[51][13] = "X"
          end
          unless var_comentario_localizacion[1][1] == nil
            hoja_trabajo.rows[51][20] = var_comentario_localizacion[1][1].join
          end
          unless var_comentario_medicion[1][1] == nil
            hoja_trabajo.rows[51][36] = "#{var_comentario_medicion[1][1].first.to_f} #{var_ud[1][1].join}"
          end
        #Deterioro3
          unless var_detalle_uno[1][2] == [nil]
            hoja_trabajo.rows[52][4] = "X"
          end
          unless var_detalle_dos[1][2] == [nil]
          hoja_trabajo.rows[52][13] = "X"
          end
          unless var_comentario_localizacion[1][2] == nil
            hoja_trabajo.rows[52][20] = var_comentario_localizacion[1][2].join
          end
          unless var_comentario_medicion[1][2] == nil
            hoja_trabajo.rows[52][36] = "#{var_comentario_medicion[1][2].first.to_f} #{var_ud[1][2].join}"
          end
        #Deterioro4
          unless var_detalle_uno[1][3] == [nil]
            hoja_trabajo.rows[53][4] = "X"
          end
          unless var_detalle_dos[1][3] == [nil]
          hoja_trabajo.rows[53][13] = "X"
          end
          unless var_comentario_localizacion[1][3] == nil
            hoja_trabajo.rows[53][20] = var_comentario_localizacion[1][3].join
          end
          unless var_comentario_medicion[1][3] == nil
            hoja_trabajo.rows[53][36] = "#{var_comentario_medicion[1][3].first.to_f} #{var_ud[1][3].join}"
          end
        #Deterioro5
          unless var_detalle_uno[1][4] == [nil]
            hoja_trabajo.rows[54][4] = "X"
          end
          unless var_comentario_localizacion[1][4] == nil
            hoja_trabajo.rows[54][20] = var_comentario_localizacion[1][4].join
          end
          unless var_comentario_medicion[1][4] == nil
            hoja_trabajo.rows[54][36] = "#{var_comentario_medicion[1][4].first.to_f} #{var_ud[1][4].join}"
          end
        #Deterioro6
          unless var_detalle_uno[1][5] == [nil]
            hoja_trabajo.rows[55][4] = "X"
          end
          unless var_comentario_localizacion[1][5] == nil
            hoja_trabajo.rows[55][20] = var_comentario_localizacion[1][5].join
          end
          unless var_comentario_medicion[1][5] == nil
            hoja_trabajo.rows[55][36] = "#{var_comentario_medicion[1][5].first.to_f} #{var_ud[1][5].join}"
          end
        #Deterioro7
          unless var_detalle_uno[1][6] == [nil]
            hoja_trabajo.rows[56][4] = "X"
          end
          unless var_comentario_localizacion[1][6] == nil
            hoja_trabajo.rows[56][20] = var_comentario_localizacion[1][6].join
          end
          unless var_comentario_medicion[1][6] == nil
            hoja_trabajo.rows[56][36] = "#{var_comentario_medicion[1][6].first.to_f} #{var_ud[1][6].join}"
          end
          #Deterioro8
            unless var_detalle_uno[1][7] == [nil]
              hoja_trabajo.rows[57][4] = "X"
            end
            unless var_comentario_localizacion[1][7] == nil
              hoja_trabajo.rows[57][20] = var_comentario_localizacion[1][7].try(:join)
            end
            unless var_comentario_medicion[1][7] == nil
              hoja_trabajo.rows[57][36] = "#{var_comentario_medicion[1][7].first.to_f} #{var_ud[1][7].join}" rescue nil
            end
      #MESOESTRUTURA
        #Nota tecnica
        if var_nota_tecnica[2] == "4"
          hoja_trabajo.rows[59][17] = 1
        elsif var_nota_tecnica[2] == "3"
          hoja_trabajo.rows[59][17] = 2
        elsif var_nota_tecnica[2] == "2"
          hoja_trabajo.rows[59][17] = 3
        elsif var_nota_tecnica[2] == "1"
          hoja_trabajo.rows[59][17] = 4
        elsif var_nota_tecnica[2] == "0"
          hoja_trabajo.rows[59][17] = 5
        end
        #Deterioro2
          unless var_detalle_uno[2][0] == [nil]
            hoja_trabajo.rows[60][4] = "X"
          end
          unless var_detalle_dos[2][0] == [nil]
          hoja_trabajo.rows[60][13] = "X"
          end
          unless var_comentario_localizacion[2][0] == nil
            hoja_trabajo.rows[60][20] = var_comentario_localizacion[2][0].join
          end
          unless var_comentario_medicion[2][0] == nil
            hoja_trabajo.rows[60][36] = "#{var_comentario_medicion[2][0].first.to_f} #{var_ud[2][0].join}"
          end
        #Deterioro2
          unless var_detalle_uno[2][1] == [nil]
            hoja_trabajo.rows[61][4] = "X"
          end
          unless var_detalle_dos[2][1] == [nil]
          hoja_trabajo.rows[61][13] = "X"
          end
          unless var_comentario_localizacion[2][1] == nil
            hoja_trabajo.rows[61][20] = var_comentario_localizacion[2][1].join
          end
          unless var_comentario_medicion[2][1] == nil
            hoja_trabajo.rows[61][36] = "#{var_comentario_medicion[2][1].first.to_f} #{var_ud[2][1].join}"
          end
        #Deterioro3
          unless var_detalle_uno[2][2] == [nil]
            hoja_trabajo.rows[62][4] = "X"
          end
          unless var_detalle_dos[2][2] == [nil]
          hoja_trabajo.rows[62][13] = "X"
          end
          unless var_comentario_localizacion[2][2] == nil
            hoja_trabajo.rows[62][20] = var_comentario_localizacion[2][2].join
          end
          unless var_comentario_medicion[2][2] == nil
            hoja_trabajo.rows[62][36] = "#{var_comentario_medicion[2][2].first.to_f} #{var_ud[2][2].join}"
          end
        #Deterioro4
          unless var_detalle_uno[2][3] == [nil]
            hoja_trabajo.rows[63][4] = "X"
          end
          unless var_detalle_dos[2][3] == [nil]
          hoja_trabajo.rows[63][13] = "X"
          end
          unless var_comentario_localizacion[2][3] == nil
            hoja_trabajo.rows[63][20] = var_comentario_localizacion[2][3].join
          end
          unless var_comentario_medicion[2][3] == nil
            hoja_trabajo.rows[63][36] = "#{var_comentario_medicion[2][3].first.to_f} #{var_ud[2][3].join}"
          end
        #Deterioro5
          unless var_detalle_uno[2][4] == [nil]
            hoja_trabajo.rows[64][4] = "X"
          end
          unless var_comentario_localizacion[2][4] == nil
            hoja_trabajo.rows[64][20] = var_comentario_localizacion[2][4].join
          end
          unless var_comentario_medicion[2][4] == nil
            hoja_trabajo.rows[64][36] = "#{var_comentario_medicion[2][4].first.to_f} #{var_ud[2][4].join}"
          end
        #Deterioro6
          unless var_detalle_uno[2][5] == [nil]
            hoja_trabajo.rows[65][4] = "X"
          end
          unless var_comentario_localizacion[2][5] == nil
            hoja_trabajo.rows[65][20] = var_comentario_localizacion[2][5].join
          end
          unless var_comentario_medicion[2][5] == nil
            hoja_trabajo.rows[65][36] = "#{var_comentario_medicion[2][5].first.to_f} #{var_ud[2][5].join}"
          end
        #Deterioro7
          unless var_detalle_uno[2][6] == [nil]
            hoja_trabajo.rows[66][4] = "X"
          end
          unless var_comentario_localizacion[2][6] == nil
            hoja_trabajo.rows[66][20] = var_comentario_localizacion[2][6].join
          end
          unless var_comentario_medicion[2][6] == nil
            hoja_trabajo.rows[66][36] = "#{var_comentario_medicion[2][6].first.to_f} #{var_ud[2][6].join}"
          end
        #Deterioro8
          unless var_detalle_uno[2][7] == [nil]
            hoja_trabajo.rows[67][4] = "X"
          end
          unless var_comentario_localizacion[2][7] == nil
            hoja_trabajo.rows[67][20] = var_comentario_localizacion[2][7].join
          end
          unless var_comentario_medicion[2][7] == nil
            hoja_trabajo.rows[67][36] = "#{var_comentario_medicion[2][7].first.to_f} #{var_ud[2][7].join}"
          end
      #INFRAESTRUTURA
        #Nota tecnica
        if var_nota_tecnica[3] == "4"
          hoja_trabajo.rows[69][17] = 1
        elsif var_nota_tecnica[3] == "3"
          hoja_trabajo.rows[69][17] = 2
        elsif var_nota_tecnica[3] == "2"
          hoja_trabajo.rows[69][17] = 3
        elsif var_nota_tecnica[3] == "1"
          hoja_trabajo.rows[69][17] = 4
        elsif var_nota_tecnica[3] == "0"
          hoja_trabajo.rows[69][17] = 5
        end
        #Deterioro3
          unless var_detalle_uno[3][0] == [nil]
            hoja_trabajo.rows[70][4] = "X"
          end
          unless var_comentario_localizacion[3][0] == nil
            hoja_trabajo.rows[70][20] = var_comentario_localizacion[3][0].join
          end
          unless var_comentario_medicion[3][0] == nil
            hoja_trabajo.rows[70][36] = "#{var_comentario_medicion[3][0].first.to_f} #{var_ud[3][0].join}"
          end
        #Deterioro3
          unless var_detalle_uno[3][1] == [nil]
            hoja_trabajo.rows[71][4] = "X"
          end
          unless var_comentario_localizacion[3][1] == nil
            hoja_trabajo.rows[71][20] = var_comentario_localizacion[3][1].join
          end
          unless var_comentario_medicion[3][1] == nil
            hoja_trabajo.rows[71][36] = "#{var_comentario_medicion[3][1].first.to_f} #{var_ud[3][1].join}"
          end
        #Deterioro3
          unless var_detalle_uno[3][2] == [nil]
            hoja_trabajo.rows[72][4] = "X"
          end
          unless var_comentario_localizacion[3][2] == nil
            hoja_trabajo.rows[72][20] = var_comentario_localizacion[3][2].join
          end
          unless var_comentario_medicion[3][2] == nil
            hoja_trabajo.rows[72][36] = "#{var_comentario_medicion[3][2].first.to_f} #{var_ud[3][2].join}"
          end
        #Deterioro4
          unless var_detalle_uno[3][3] == [nil]
            hoja_trabajo.rows[73][4] = "X"
          end
          unless var_comentario_localizacion[3][3] == nil
            hoja_trabajo.rows[73][20] = var_comentario_localizacion[3][3].join
          end
          unless var_comentario_medicion[3][3] == nil
            hoja_trabajo.rows[73][36] = "#{var_comentario_medicion[3][3].first.to_f} #{var_ud[3][3].join}"
          end
      #PISTA/ACCESO
        #Nota tecnica
        # byebug
        if var_nota_tecnica[4] == "4"
          hoja_trabajo.rows[75][17] = 1
        elsif var_nota_tecnica[4] == "3"
          hoja_trabajo.rows[75][17] = 2
        elsif var_nota_tecnica[4] == "2"
          hoja_trabajo.rows[75][17] = 3
        elsif var_nota_tecnica[4] == "1"
          hoja_trabajo.rows[75][17] = 4
        elsif var_nota_tecnica[4] == "0"
          hoja_trabajo.rows[75][17] = 5
        end
        #Deterioro4
          unless var_detalle_uno[4][0] == [nil]
            hoja_trabajo.rows[76][4] = "X"
          end
          unless var_detalle_dos[4][0] == [nil]
          hoja_trabajo.rows[76][13] = "X"
          end
          unless var_comentario_localizacion[4][0] == nil
            hoja_trabajo.rows[76][21] = var_comentario_localizacion[4][0].join
          end
          unless var_comentario_medicion[4][0] == nil
            hoja_trabajo.rows[76][36] = "#{var_comentario_medicion[4][0].first.to_f} #{var_ud[4][0].join}"
          end
        #Deterioro4
          unless var_detalle_uno[4][1] == [nil]
            hoja_trabajo.rows[77][4] = "X"
          end
          unless var_detalle_dos[4][1] == [nil]
          hoja_trabajo.rows[77][13] = "X"
          end
          unless var_comentario_localizacion[4][1] == nil
            hoja_trabajo.rows[77][21] = var_comentario_localizacion[4][1].join
          end
          unless var_comentario_medicion[4][1] == nil
            hoja_trabajo.rows[77][36] = "#{var_comentario_medicion[4][1].first.to_f} #{var_ud[4][1].join}"
          end
        #Deterioro4
          unless var_detalle_uno[4][2] == [nil]
            hoja_trabajo.rows[78][4] = "X"
          end
          unless var_detalle_dos[4][2] == [nil]
          hoja_trabajo.rows[78][13] = "X"
          end
          unless var_comentario_localizacion[4][2] == nil
            hoja_trabajo.rows[78][21] = var_comentario_localizacion[4][2].join
          end
          unless var_comentario_medicion[4][2] == nil
            hoja_trabajo.rows[78][36] = "#{var_comentario_medicion[4][2].first.to_f} #{var_ud[4][2].join}"
          end
        #Deterioro4
          unless var_detalle_uno[4][3] == [nil]
            hoja_trabajo.rows[79][4] = "X"
          end
          unless var_detalle_dos[4][3] == [nil]
          hoja_trabajo.rows[79][13] = "X"
          end
          unless var_comentario_localizacion[4][3] == nil
            hoja_trabajo.rows[79][21] = var_comentario_localizacion[4][3].join
          end
          unless var_comentario_medicion[4][3] == nil
            hoja_trabajo.rows[79][36] = "#{var_comentario_medicion[4][3].first.to_f} #{var_ud[4][3].join}"
          end
      #PARAMETROS DE DESEMPENHO
        #Deterioro1
        unless var_detalle_uno[5][0] == [nil]
          hoja_trabajo.rows[82][9] = "X"
        end
        unless var_detalle_dos[5][0] == [nil]
          hoja_trabajo.rows[82][18] = "X"
        end
        unless var_comentario_observaciones[5][0] == [nil] || var_comentario_observaciones[5][0] == nil
          hoja_trabajo.rows[82][25] = var_comentario_observaciones[5][0].join
        end
        #Deterioro2
        unless var_detalle_uno[5][1] == [nil]
          hoja_trabajo.rows[83][9] = "X"
        end
        unless var_detalle_dos[5][1] == [nil]
          hoja_trabajo.rows[83][18] = "X"
        end
        unless var_comentario_observaciones[5][1] == [nil] || var_comentario_observaciones[5][1] == nil
          hoja_trabajo.rows[83][25] = var_comentario_observaciones[5][1].join
        end
        #Deterioro3
        unless var_detalle_uno[5][2] == [nil]
          hoja_trabajo.rows[84][9] = "X"
        end
        unless var_detalle_dos[5][2] == [nil]
          hoja_trabajo.rows[84][18] = "X"
        end
        unless var_comentario_observaciones[5][2] == [nil] || var_comentario_observaciones[5][2] == nil
          hoja_trabajo.rows[84][25] = var_comentario_observaciones[5][2].join
        end
        #Deterioro4
        unless var_detalle_uno[5][3] == [nil]
          hoja_trabajo.rows[85][9] = "X"
        end
        unless var_detalle_dos[5][3] == [nil]
          hoja_trabajo.rows[85][18] = "X"
        end
        unless var_detalle_tres[5][2] == [nil]
          hoja_trabajo.rows[85][18] = "X"
        end
        unless var_comentario_observaciones[5][3] == [nil] || var_comentario_observaciones[5][3] == nil
          hoja_trabajo.rows[85][25] = var_comentario_observaciones[5][3].join
        end
        #Deterioro5
        unless var_detalle_uno[5][4] == [nil]
          hoja_trabajo.rows[86][9] = "X"
        end
        unless var_detalle_dos[5][4] == [nil]
          hoja_trabajo.rows[86][18] = "X"
        end
        unless var_comentario_observaciones[5][4] == nil || var_comentario_observaciones[5][4] == nil
          hoja_trabajo.rows[86][25] = var_comentario_observaciones[5][4].join
        end

        Create_excel.new_excel_images(id_json)

        open_book.write('/tmp/arteris_excel.xls')
        #convertir a excel y descargar en prod
        system('soffice --headless --calc --infilter="Microsoft Excel 97/2000/XP" --convert-to xlsx:"Calc MS Excel 2007 XML" /tmp/arteris_excel.xls --outdir /tmp/')
        
        send_file "/tmp/arteris_excel.xlsx", :filename => "arteris_excel.xlsx", :type => 'application/xlsx'
        #descargar en local en xls
        #send_file "/tmp/arteris_excel.xls", :filename => "arteris_excel.xls", :type => 'application/xls'
        #descargar fotos arteris
  end


  def download_attachments
    id = params[:id]
    controller = params[:controller]
    model = singularize_me(controller).camelize.constantize
    arr = Array.new
    if model.to_s == "InspeccionBasica"
      @inspeccion = model.includes(:subinformes => {:elementos => :adjuntos}).where(:ib_id => id).first

      @inspeccion.subinformes.each do |sub|
        sub.elementos.each do |adj|
          arr << adj.adjuntos.map { |x| x.id } if !adj.adjuntos.empty?
        end
      end
    elsif model.to_s == "InspeccionPrincipal"
      @inspeccion = model.includes(:elementos => :adjuntos).where(:ip_id => id).first

      @inspeccion.elementos.each do |adj|
        arr << adj.adjuntos.map { |x| x.id } if !adj.adjuntos.empty?
      end
    elsif model.to_s == "InspeccionEspecial"
      @inspeccion = model.includes(:elementos => :adjuntos).where(:ip_ie => id).first

      @inspeccion.elementos.each do |adj|
        arr << adj.adjuntos.map { |x| x.id } if !adj.adjuntos.empty?
      end
    else
      @inspeccion = model.includes(:elementos => :adjuntos).where(:it_id => id).first
    end


    arr << @inspeccion.adjuntos.map { |x| x.id } if !@inspeccion.adjuntos.empty?
    arr = arr.flatten.to_set.to_a

    attachments = Attachment.where(id: arr).all
    generate_zip_with_attachments(attachments)
  end

  def generate_zip_with_attachments(attachments)
    zipfile = Tempfile.new('incabridges')
    path = zipfile.path
    zipfile.unlink
    Zip::File.open(path, Zip::File::CREATE) do |zipfile|
      attachments.each do |att|
        begin
          zipfile.add(att.name, att.app_path)
        rescue => e
          Rails.logger.warn e
        end
      end
    end

    send_file path, :filename => @inspeccion.id.to_s+"_"+DateTime.now.strftime('%Y%m%d-%H%M')+'.zip', :type => 'application/zip'
  end

  private

  # Nombre del fichero del informe relacionado a una inspeccion. Por defecto
  # es el codigo de la estructura.
  #
  # @param [Actuacion] actuacion objeto al que corresponde el informe generado
  # @param [Symbol] _tipo tipo de informe
  #   con el tipo del informe (:inspeccion)
  # @return [String] cadena con el nombre correspondiente al informe
  def nombre_informe(inspeccion, _tipo)
    inspeccion.inventario.inv_codigopuente
  end

  def inspection_filter_report(inspeccion)
    return nil unless SETTINGS.filter_inspection_report_enabled?
    conditions = INSPECTION_FILTER_FIELDS.map do |md|
      if params[md.attribute_name].present?
        {md.attribute_name => params[md.attribute_name]}
      else
        {}
      end
    end.inject({}) { |hash, injected| hash.merge!(injected) }
    filtered_elements = inspeccion.elementos.where(conditions)
    if params[:min_pos].present? && params[:max_pos].present?
      filtered_elements = filtered_elements.check_posicion(params[:min_pos],
                                                           params[:max_pos])
    end
    filtered_elements
  end
end
