import flet as ft
import openpyxl
from openpyxl.drawing.image import Image
import os
import shutil
from datetime import datetime

def main(page: ft.Page):
    page.title = "Tekstil Denetim Uygulaması"
    page.scroll = "adaptive"
    page.theme_mode = ft.ThemeMode.DARK
    
    # Resim yollarını tutacak sözlük
    secilen_resimler = {
        "paketli_on": None,
        "paketli_arka": None,
        "sirali_acik": None,
        "etiket_yikama": None,
        "etiket_kart": None,
        "olcum_1": None,
        "olcum_2": None,
        "olcum_3": None,
        "defo_1": None,
        "defo_2": None,
    }

    # Excel hücre eşleşmeleri ve başlıkları
    resim_ayarlari = {
        "paketli_on": {"hucre": "A40", "baslik": "PAKETLİ ÖN GÖRÜNÜM"},
        "paketli_arka": {"hucre": "I40", "baslik": "PAKETLİ ARKA GÖRÜNÜM"},
        "sirali_acik": {"hucre": "A57", "baslik": "SIRALI AÇIK GÖRÜNÜM"},
        "etiket_yikama": {"hucre": "I57", "baslik": "YIKAMA TALİMATI ETİKETİ"},
        "etiket_kart": {"hucre": "A74", "baslik": "KART ETİKETİ"},
        "olcum_1": {"hucre": "I74", "baslik": "ÖLÇÜM TABLOSU REHBERİ"},
        "olcum_2": {"hucre": "A91", "baslik": "ÖLÇÜM DETAYI 1"},
        "olcum_3": {"hucre": "I91", "baslik": "ÖLÇÜM DETAYI 2"},
        "defo_1": {"hucre": "A108", "baslik": "HATA / DEFO DETAYI 1"},
        "defo_2": {"hucre": "I108", "baslik": "HATA / DEFO DETAYI 2"},
    }

    durum_metni = ft.Text("", color=ft.colors.GREEN_ACCENT, weight=ft.FontWeight.BOLD)
    
    # Android dosya yolları
    sablon_adi = "INSPECTION REPORT.xlsx"
    indirilenler_klasoru = "/storage/emulated/0/Download"
    sablon_yolu = os.path.join(indirilenler_klasoru, sablon_adi)
    cikis_yolu = os.path.join(indirilenler_klasoru, "TAMAMLANMIS_RAPOR.xlsx")

    # Resim Seçme İşlemi
    def dosya_secildi(e: ft.FilePickerResultEvent):
        if e.files:
            secilen_yol = e.files[0].path
            buton_anahtari = e.control.data
            secilen_resimler[buton_anahtari] = secilen_yol
            
            # Arayüzdeki ilgili butonu bulup rengini değiştirme
            for satir in resim_listesi.controls:
                if isinstance(satir, ft.Row):
                    for b in satir.controls:
                        if isinstance(b, ft.ElevatedButton) and b.data == buton_anahtari:
                            b.bgcolor = ft.colors.GREEN_700
                            b.text = f"✓ {resim_ayarlari[buton_anahtari]['baslik']} Seçildi"
            page.update()

    dosya_secici = ft.FilePicker(on_result=dosya_secildi)
    page.overlay.append(dosya_secici)

    def resim_secici_butonu(baslik, anahtar):
        return ft.ElevatedButton(
            text=baslik,
            icon=ft.icons.CAMERA_ALT,
            width=280,
            height=50,
            data=anahtar,
            on_click=lambda e: dosya_secici.pick_files(allow_multiple=False, file_type=ft.FilePickerFileType.IMAGE)
        )

    # Excel Kaydetme Fonksiyonu
    def raporu_olustur(e):
        if not os.path.exists(sablon_yolu):
            durum_metni.value = "HATA: Download klasöründe 'INSPECTION REPORT.xlsx' bulunamadı!"
            durum_metni.color = ft.colors.RED_ACCENT
            page.update()
            return

        try:
            durum_metni.value = "Rapor hazırlanıyor, lütfen bekleyin..."
            durum_metni.color = ft.colors.BLUE_ACCENT
            page.update()

            wb = openpyxl.load_workbook(sablon_yolu)
            ws = wb.active

            for anahtar, yol in secilen_resimler.items():
                if yol and os.path.exists(yol):
                    hucre = resim_ayarlari[anahtar]["hucre"]
                    img = Image(yol)
                    # Hücre boyutuna göre resmi otomatik ölçekleme (Yaklaşık 340x310 piksel)
                    img.width = 340
                    img.height = 310
                    ws.add_image(img, hucre)

            wb.save(cikis_yolu)
            durum_metni.value = f"BAŞARILI! Rapor 'Download/TAMAMLANMIS_RAPOR.xlsx' olarak kaydedildi."
            durum_metni.color = ft.colors.GREEN_ACCENT
            
        except Exception as ex:
            durum_metni.value = f"Hata Oluştu: {str(ex)}"
            durum_metni.color = ft.colors.RED_ACCENT
        
        page.update()

    # Arayüz Bileşenleri
    resim_listesi = ft.Column(
        controls=[
            ft.Row([resim_secici_butonu("Paketli Ön Görünüm", "paketli_on"), resim_secici_butonu("Paketli Arka Görünüm", "paketli_arka")], alignment=ft.MainAxisAlignment.CENTER),
            ft.Row([resim_secici_butonu("Sıralı Açık Görünüm", "sirali_acik"), resim_secici_butonu("Yıkama Talimatı", "etiket_yikama")], alignment=ft.MainAxisAlignment.CENTER),
            ft.Row([resim_secici_butonu("Kart Etiketi", "etiket_kart"), resim_secici_butonu("Ölçüm Rehberi", "olcum_1")], alignment=ft.MainAxisAlignment.CENTER),
            ft.Row([resim_secici_butonu("Ölçüm Detayı 1", "olcum_2"), resim_secici_butonu("Ölçüm Detayı 2", "olcum_3")], alignment=ft.MainAxisAlignment.CENTER),
            ft.Row([resim_secici_butonu("Hata / Defo 1", "defo_1"), resim_secici_butonu("Hata / Defo 2", "defo_2")], alignment=ft.MainAxisAlignment.CENTER),
        ],
        spacing=15,
        horizontal_alignment=ft.CrossAxisAlignment.CENTER
    )

    page.add(
        ft.Container(
            content=ft.Column([
                ft.Text("TEKSTİL DENETİM RAPORLAMA", size=24, weight=ft.FontWeight.BOLD, color=ft.colors.BLUE_200),
                ft.Divider(height=20, color=ft.colors.WHITE24),
                resim_listesi,
                ft.Divider(height=30, color=ft.colors.WHITE24),
                ft.ElevatedButton(
                    text="EXCEL RAPORUNU OLUŞTUR",
                    icon=ft.icons.REFLOW,
                    bgcolor=ft.colors.BLUE_700,
                    color=ft.colors.WHITE,
                    width=350,
                    height=60,
                    on_click=raporu_olustur
                ),
                ft.Container(content=durum_metni, margin=ft.margin.only(top=20))
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
            alignment=ft.alignment.center,
            padding=20
        )
    )

ft.app(target=main)