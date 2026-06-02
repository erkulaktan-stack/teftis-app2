import flet as ft
import openpyxl
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
import os

def main(page: ft.Page):
    page.title = "ÖZDEMİRTEKS Tablet Teftiş Paneli"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.scroll = ft.ScrollMode.AUTO
    
    page.window_width = 850
    page.window_height = 850
    
    # Android'in her uygulamaya koşulsuz şartsız izin verdiği iç depolama klasörü
    uygulama_klasoru = os.getcwd()
    rapor_yolu = os.path.join(uygulama_klasoru, "DENETIM_RAPORU_SABLON.xlsx")

    # --- SIFIRDAN EXCEL OLUŞTURMA MOTORU ---
    # Eğer içeride şablon yoksa, uygulama hata vermek yerine sekmeleri kendi oluşturur!
    if not os.path.exists(rapor_yolu):
        wb = Workbook()
        # 1. Sekme
        ws1 = wb.active
        ws1.title = "AQL - PAPERWORK COPY"
        # Başlıkları dolduruyoruz ki şablon gibi dursun
        ws1["A1"] = "ÖZDEMİRTEKS DENETİM RAPORU"
        
        # 2. Sekme
        wb.create_sheet(title="AQL MEASUREMENT CHART COPY")
        
        # 3. Sekme
        wb.create_sheet(title="Photos - AQL COPY")
        
        wb.save(rapor_yolu)

    # --- SAYAÇLAR ---
    sayaclar = {"major": 0, "minor": 0}
    major_txt = ft.Text(value="0", size=28, weight=ft.FontWeight.BOLD, color=ft.Colors.RED_800)
    minor_txt = ft.Text(value="0", size=28, weight=ft.FontWeight.BOLD, color=ft.Colors.ORANGE_800)

    def sayac_degistir(islem, tip):
        if islem == "artir":
            sayaclar[tip] += 1
        elif islem == "azalt" and sayaclar[tip] > 0:
            sayaclar[tip] -= 1
        major_txt.value = str(sayaclar["major"])
        minor_txt.value = str(sayaclar["minor"])
        page.update()

    # --- RESİM SEÇİM SÖZLÜKLERİ ---
    secilen_resimler = {
        "olcu_xs": None, "olcu_s": None, "olcu_m": None, "olcu_l": None, "olcu_xl": None, "olcu_xxl": None,
        "minor_img": None, "major_img": None, "onden": None, "arkadan": None, "paketli_on": None, "paketli_arka": None,
    }

    onizleme_yazilari = {k: ft.Text("Seçilmedi", size=12, italic=True, color=ft.Colors.GREY_600) for k in secilen_resimler.keys()}
    aktif_kategori = [None]

    def dosya_secildi_sonuc(e: ft.FilePickerResultEvent):
        if e.files and aktif_kategori[0]:
            dosya_yolu = e.files[0].path
            kat = aktif_kategori[0]
            secilen_resimler[kat] = dosya_yolu
            onizleme_yazilari[kat].value = os.path.basename(dosya_yolu)
            onizleme_yazilari[kat].color = ft.Colors.GREEN_700
            onizleme_yazilari[kat].weight = ft.FontWeight.BOLD
            page.update()

    file_picker = ft.FilePicker(on_result=dosya_secildi_sonuc)
    page.overlay.append(file_picker)

    def tetikle_dosya_sec(kategori):
        aktif_kategori[0] = kategori
        file_picker.pick_files(allow_multiple=False, file_type=ft.FilePickerFileType.IMAGE)

    def resim_secici_butonu(etiket, kategori):
        return ft.Row([
            ft.Button(
                content=ft.Row([
                    ft.Icon(ft.Icons.CAMERA_ALT, color=ft.Colors.BLACK),
                    ft.Text(value=etiket, color=ft.Colors.BLACK, weight=ft.FontWeight.W_500)
                ], alignment=ft.MainAxisAlignment.CENTER, spacing=5),
                on_click=lambda e: tetikle_dosya_sec(kategori),
                width=260,
                style=ft.ButtonStyle(bgcolor=ft.Colors.BLUE_GREY_100)
            ),
            onizleme_yazilari[kategori]
        ], alignment=ft.MainAxisAlignment.START, spacing=10)

    # --- TEXT INPUT ALANLARI ---
    brand_dropdown = ft.Dropdown(label="Marka (Brand)", options=[ft.dropdown.Option("JD SPORT"), ft.dropdown.Option("Gymking")], value="JD SPORT", width=330)
    supplier_input = ft.TextField(label="Tedarikçi (Supplier)", value="OZDEMIRTEKS TEKSTIL", width=330)
    factory_input = ft.TextField(label="Fabrika (Factory)", value="OZDEMIRTEKS TEKSTIL", width=330)
    inspector_input = ft.TextField(label="Denetçi (Checked by)", value="Merdali Mete", width=330)
    po_input = ft.TextField(label="PO Numarası", value="123456", width=330)
    style_name_input = ft.TextField(label="Stil Adı (Style Name)", value="JD SPORT", width=330)
    style_num_input = ft.TextField(label="Stil No (Style Number)", value="none", width=330)
    order_qty_input = ft.TextField(label="Sipariş Miktarı", value="2688", width=330)
    colour_input = ft.TextField(label="Renk (Colour)", value="Siyah", width=330)

    form_sol = ft.Column([brand_dropdown, supplier_input, factory_input, inspector_input, colour_input])
    form_sag = ft.Column([po_input, style_name_input, style_num_input, order_qty_input])

    def ekrani_degistir(e):
        ekran_1.visible = False
        ekran_2.visible = True
        page.update()

    def ana_ekrana_don(e):
        ekran_2.visible = False
        ekran_1.visible = True
        page.update()

    buton_sonraki = ft.Button(
        content=ft.Text("Sonraki Adım: Resim ve Hata Girişi", color=ft.Colors.WHITE, weight=ft.FontWeight.BOLD), 
        bgcolor=ft.Colors.BLUE_800, width=320, height=50, on_click=ekrani_degistir
    )

    ekran_1 = ft.Column(controls=[
        ft.Row([ft.Text("ÖZDEMİRTEKS - INSPECTION PROGRAMI", size=20, weight=ft.FontWeight.BOLD)]),
        ft.Divider(), 
        ft.Row([form_sol, form_sag], spacing=20), 
        ft.Divider(),
        ft.Row([buton_sonraki], alignment=ft.MainAxisAlignment.CENTER)
    ], visible=True)

    bilgi_notu = ft.Text(value="", size=14, weight=ft.FontWeight.BOLD, color=ft.Colors.GREEN_700)

    def excele_raporu_yaz(e):
        # Çıktıyı uygulamanın kendi izinli klasörüne doğrudan kaydeder
        cikti_yolu = os.path.join(uygulama_klasoru, f"TEFTIS_RAPORU_PO_{po_input.value}.xlsx")
        
        try:
            bilgi_notu.value = "Rapor Excel'e yazılıyor..."
            page.update()

            wb = load_workbook(rapor_yolu)
            ws1 = wb["AQL - PAPERWORK COPY"]
            
            # Form verilerini hücrelere yazma
            ws1["A10"] = f"PO: {po_input.value}"
            ws1["A11"] = f"STYLE NAME: {style_name_input.value}"
            ws1["A12"] = f"STYLE NUMBER: {style_num_input.value}"
            ws1["A13"] = f"SUPPLIER: {supplier_input.value}"
            ws1["A14"] = f"FACTORY: {factory_input.value}"
            ws1["A15"] = f"Checked by: {inspector_input.value}"
            ws1["E9"] = sayaclar["major"]
            ws1["F9"] = sayaclar["minor"]

            # Resimleri Ekleme Haritası
            resim_haritasi = [
                ("olcu_xs", "AQL MEASUREMENT CHART COPY", "A7"),
                ("olcu_s", "AQL MEASUREMENT CHART COPY", "P7"),
                ("minor_img", "Photos - AQL COPY", "C3"),
                ("major_img", "Photos - AQL COPY", "I3"),
                ("onden", "Photos - AQL COPY", "B20"),
                ("arkadan", "Photos - AQL COPY", "C20"),
                ("paketli_on", "Photos - AQL COPY", "I20"),
                ("paketli_arka", "Photos - AQL COPY", "O20"),
            ]

            for anahtar, sekme_adi, hucre in resim_haritasi:
                resim_yolu = secilen_resimler[anahtar]
                if resim_yolu and os.path.exists(resim_yolu):
                    ws_resim = wb[sekme_adi]
                    img = OpenpyxlImage(resim_yolu)
                    img.width, img.height = (240, 320) if "MEASUREMENT" in sekme_adi else (280, 210)
                    ws_resim.add_image(img, hucre)

            wb.save(cikti_yolu)
            bilgi_notu.value = f"Başarılı! Rapor kaydedildi.\nDosya: {cikti_yolu}"
            bilgi_notu.color = ft.Colors.GREEN_700
        except Exception as ex:
            bilgi_notu.value = f"Hata: {str(ex)}"
            bilgi_notu.color = ft.Colors.RED_700
        page.update()

    # --- SAYAÇ VE RESİM PANELİ ---
    sayac_paneli = ft.Row([
        ft.Container(content=ft.Column([
            ft.Text("MAJOR HATA", size=14, weight=ft.FontWeight.BOLD),
            ft.Row([
                ft.IconButton(ft.Icons.REMOVE_CIRCLE, on_click=lambda e: sayac_degistir("azalt", "major"), icon_color=ft.Colors.RED_700),
                major_txt,
                ft.IconButton(ft.Icons.ADD_CIRCLE, on_click=lambda e: sayac_degistir("artir", "major"), icon_color=ft.Colors.RED_700)
            ], alignment=ft.MainAxisAlignment.CENTER)
        ], horizontal_alignment=ft.CrossAxisAlignment.CENTER), padding=10, border=ft.border.all(1, ft.Colors.RED_200), border_radius=8, width=200),
        
        ft.Container(content=ft.Column([
            ft.Text("MINOR HATA", size=14, weight=ft.FontWeight.BOLD),
            ft.Row([
                ft.IconButton(ft.Icons.REMOVE_CIRCLE, on_click=lambda e: sayac_degistir("azalt", "minor"), icon_color=ft.Colors.ORANGE_700),
                minor_txt,
                ft.IconButton(ft.Icons.ADD_CIRCLE, on_click=lambda e: sayac_degistir("artir", "minor"), icon_color=ft.Colors.ORANGE_700)
            ], alignment=ft.MainAxisAlignment.CENTER)
        ], horizontal_alignment=ft.CrossAxisAlignment.CENTER), padding=10, border=ft.border.all(1, ft.Colors.ORANGE_200), border_radius=8, width=200)
    ], spacing=30, alignment=ft.MainAxisAlignment.CENTER)

    resim_listesi = ft.Column([
        ft.Text("Ölçüm Tablosu Fotoğrafları", weight=ft.FontWeight.BOLD, size=14),
        resim_secici_butonu("XS Ölçüm Kartı", "olcu_xs"),
        resim_secici_butonu("S Ölçüm Kartı", "olcu_s"),
        ft.Divider(),
        ft.Text("Genel Teftiş Fotoğrafları", weight=ft.FontWeight.BOLD, size=14),
        resim_secici_butonu("Minor Hata Görseli", "minor_img"),
        resim_secici_butonu("Major Hata Görseli", "major_img"),
        resim_secici_butonu("Ürün Önden Görünüm", "onden"),
        resim_secici_butonu("Ürün Arkadan Görünüm", "arkadan"),
        resim_secici_butonu("Paketli Ön Görünüm", "paketli_on"),
        resim_secici_butonu("Paketli Arka Görünüm", "paketli_arka"),
    ], spacing=10)

    ekran_2 = ft.Column(controls=[
        ft.Row([ft.IconButton(ft.Icons.ARROW_BACK, on_click=ana_ekrana_don), ft.Text("Hata Sayacı ve Resim Paneli", size=18, weight=ft.FontWeight.BOLD)]),
        ft.Divider(), sayac_paneli, ft.Divider(), resim_listesi, ft.Divider(),
        ft.Row([ft.ElevatedButton(content=ft.Text("EXCEL RAPORUNU OLUŞTUR", color=ft.Colors.WHITE, weight=ft.FontWeight.BOLD), bgcolor=ft.Colors.GREEN_700, width=350, height=55, on_click=excele_raporu_yaz)], alignment=ft.MainAxisAlignment.CENTER),
        ft.Row([bilgi_notu], alignment=ft.MainAxisAlignment.CENTER)
    ], visible=False)

    page.clean()
    page.add(ekran_1, ekran_2)
    page.update()

ft.app(target=main)
