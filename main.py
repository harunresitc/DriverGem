import sys
import os
import platform
import re
import wmi
import webbrowser
import google.generativeai as genai

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QIcon, QAction, QGuiApplication
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QLabel, QLineEdit, QStatusBar, QMessageBox
)

# --- Modern/Koyu Tema için QSS (Qt Style Sheet) ---
# Arayüzü "şık" hale getiren kısım budur.
APP_STYLESHEET = """
QWidget {
    font-size: 10pt;
    background-color: #2b2b2b;
    color: #f0f0f0;
}
QMainWindow {
    background-color: #2b2b2b;
}
QStatusBar {
    font-size: 9pt;
}
QHeaderView::section {
    background-color: #3c3c3c;
    color: #f0f0f0;
    padding: 6px;
    border: 1px solid #555;
    font-weight: bold;
}
QTableWidget {
    gridline-color: #555;
    border: 1px solid #555;
}
QTableWidget::item {
    padding: 5px;
}
QTableWidget::item:selected {
    background-color: #5a9; /* Seçim rengi */
    color: #000;
}
QPushButton {
    background-color: #0078d4;
    color: white;
    font-weight: bold;
    border: none;
    padding: 10px 15px;
    border-radius: 4px;
}
QPushButton:hover {
    background-color: #0088f0;
}
QPushButton:pressed {
    background-color: #006ac0;
}
QPushButton:disabled {
    background-color: #555;
    color: #999;
}
QLineEdit {
    background-color: #3c3c3c;
    border: 1px solid #555;
    padding: 5px;
    border-radius: 4px;
}
QLabel {
    padding-bottom: 2px;
}
"""

# --- Arka Plan İş Parçacığı (Worker Thread) ---
# Arayüzün donmasını engelleyen en kritik sınıf
class DriverFinderThread(QThread):
    """
    WMI ve Gemini sorgularını arka planda yürüten iş parçacığı.
    """
    # Sinyaller: Bu thread, ana arayüze bu sinyallerle bilgi gönderir
    status_update = Signal(str)            # Durum çubuğu mesajı için
    hardware_found = Signal(dict)          # Bulunan her donanım için (tabloya eklemek)
    link_found = Signal(int, str, str)     # (satır_no, link, hata_mesajı)
    finished = Signal()                    # İşlem bittiğinde

    def __init__(self, api_key):
        super().__init__()
        self.api_key = api_key
        self.os_info = self._get_system_os_info()
        self.model = None

    def _get_system_os_info(self):
        try:
            bitness = "64-bit" if "64" in platform.machine() else "32-bit"
            return f"{platform.system()} {platform.release()} {bitness}"
        except Exception:
            return "Windows 10 64-bit"

    def _get_link_from_gemini(self, device_info):
        """Gemini API'ye sürücü linki sorar."""
        prompt = f"""
        GÖREV: Aşağıdaki donanım için RESMİ sürücü indirme sayfasının URL'ini bul.
        DONANIM:
        - Adı: {device_info['name']}
        - Vendor ID: VEN_{device_info['ven']}
        - Device ID: DEV_{device_info['dev']}
        - İşletim Sistemi: {self.os_info}
        KURALLAR:
        1. Sadece resmi üretici (NVIDIA, Intel, AMD, Realtek vb.) veya
           bilgisayar üreticisi (Dell, HP, Lenovo) linki ver.
        2. ASLA üçüncü parti sürücü indirme sitelerinden (driverpack vb.) link verme.
        3. Sadece URL'i döndür. Ekstra açıklama ("İşte link:" vb.) ekleme.
        4. Güvenli/resmi link bulamazsan, sadece "BULUNAMADI" yaz.
        URL:
        """
        try:
            response = self.model.generate_content(prompt)
            link = response.text.strip().replace("`", "")
            return link, None
        except Exception as e:
            error_msg = str(e)
            print(f"Gemini Hatası: {error_msg}")
            if "API key not valid" in error_msg:
                return "HATA", "API Anahtarı geçersiz."
            return "HATA", "API sorgusu başarısız."

    def run(self):
        """Thread'in ana çalışma fonksiyonu."""
        try:
            genai.configure(api_key=self.api_key)
            self.model = genai.GenerativeModel('gemini-1.5-flash')
            # API'yi test etmek için küçük bir sorgu
            self.model.generate_content("test")
        except Exception as e:
            self.status_update.emit(f"HATA: API Anahtarı geçersiz veya bağlantı sorunu.")
            self.finished.emit()
            return
        
        # 1. Adım: WMI ile Donanımları Tara
        self.status_update.emit("WMI ile yerel donanımlar taranıyor...")
        ven_dev_regex = re.compile(r'VEN_([0-9A-F]{4})&DEV_([0-9A-F]{4})', re.IGNORECASE)
        hardware_list = []
        
        try:
            c = wmi.WMI()
            devices = c.Win32_PnPEntity()
            
            for device in devices:
                if device.DeviceID:
                    match = ven_dev_regex.search(device.DeviceID)
                    if match:
                        name = device.Name if device.Name else "Bilinmeyen Aygıt"
                        hw_info = {
                            "name": name,
                            "ven": match.group(1),
                            "dev": match.group(2)
                        }
                        hardware_list.append(hw_info)
                        self.hardware_found.emit(hw_info) # Arayüzü anlık güncelle
            
            if not hardware_list:
                self.status_update.emit("Taranacak donanım kimliği bulunamadı.")
                self.finished.emit()
                return

        except wmi.x_wmi:
            self.status_update.emit("HATA: WMI sorgusu başarısız. (Yönetici izni gerekli mi?)")
            self.finished.emit()
            return
        
        # 2. Adım: Her donanım için Gemini'ye sor
        self.status_update.emit(f"{len(hardware_list)} donanım için sürücüler aranıyor (Gemini API)...")
        
        for i, hw_info in enumerate(hardware_list):
            self.status_update.emit(f"Aranıyor ({i+1}/{len(hardware_list)}): {hw_info['name'][:50]}...")
            link, error = self._get_link_from_gemini(hw_info)
            self.link_found.emit(i, link, error) # Arayüzdeki satırı güncelle
            
            if error: # API hatası durumunda döngüden çık
                self.status_update.emit(f"HATA: {error}")
                break

        self.status_update.emit("Tarama tamamlandı.")
        self.finished.emit()

# --- Ana Arayüz Sınıfı (QMainWindow) ---
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.thread = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Gemini Sürücü Bulucu")
        
        # Uygulama ikonu (Qt'nin standart ikonlarından birini kullanıyoruz)
        app_icon = self.style().standardIcon(QIcon.StandardPixmap.SP_DriveNetIcon)
        self.setWindowIcon(app_icon)
        
        self.setGeometry(100, 100, 900, 600)

        # Ana widget ve layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # --- API Anahtarı Alanı ---
        api_layout = QHBoxLayout()
        api_label = QLabel("Gemini API Anahtarı:")
        self.api_key_input = QLineEdit()
        self.api_key_input.setPlaceholderText("API anahtarınızı buraya yapıştırın (örn: AIza...)")
        self.api_key_input.setEchoMode(QLineEdit.EchoMode.Password)
        
        api_layout.addWidget(api_label)
        api_layout.addWidget(self.api_key_input)
        
        # --- Başlat Butonu ---
        self.scan_button = QPushButton("Donanımları Tara ve Sürücüleri Bul")
        # İkon ekleme
        scan_icon = self.style().standardIcon(QIcon.StandardPixmap.SP_SearchQuery)
        self.scan_button.setIcon(scan_icon)
        self.scan_button.clicked.connect(self.start_scan)
        
        api_layout.addWidget(self.scan_button)
        main_layout.addLayout(api_layout)

        # --- Sonuç Tablosu ---
        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(3)
        self.table_widget.setHorizontalHeaderLabels(["Aygıt Adı", "Donanım Kimliği", "Önerilen Sürücü Linki"])
        self.table_widget.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table_widget.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table_widget.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.table_widget.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers) # Düzenlemeyi kapat
        
        # Çift tıklama ile linki kopyalama veya açma
        self.table_widget.cellDoubleClicked.connect(self.on_cell_double_clicked)
        
        main_layout.addWidget(self.table_widget)

        # --- Durum Çubuğu ---
        self.setStatusBar(QStatusBar(self))
        self.statusBar().showMessage("Hazır. Başlamak için API anahtarınızı girin ve Tara'ya tıklayın.")

        # Stili uygula
        self.setStyleSheet(APP_STYLESHEET)
    
    def start_scan(self):
        """Tarama butonuna tıklandığında çalışır."""
        api_key = self.api_key_input.text().strip()
        if not api_key:
            QMessageBox.warning(self, "API Anahtarı Eksik", 
                                "Lütfen devam etmek için Gemini API anahtarınızı girin.")
            return

        # Arayüzü kilitle
        self.scan_button.setEnabled(False)
        self.scan_button.setText("Taranıyor...")
        self.table_widget.setRowCount(0) # Tabloyu temizle
        
        # Arka plan thread'ini başlat
        self.thread = DriverFinderThread(api_key)
        
        # Sinyalleri ilgili fonksiyonlara (slot) bağla
        self.thread.status_update.connect(self.update_status)
        self.thread.hardware_found.connect(self.add_hardware_row)
        self.thread.link_found.connect(self.update_driver_link)
        self.thread.finished.connect(self.scan_finished)
        
        self.thread.start()

    # --- Sinyal Slotları (Thread'den gelen veriyi işleyen fonksiyonlar) ---
    
    def update_status(self, message):
        """Durum çubuğunu günceller."""
        self.statusBar().showMessage(message)

    def add_hardware_row(self, hw_info):
        """Tabloya yeni bir donanım satırı ekler."""
        row_count = self.table_widget.rowCount()
        self.table_widget.insertRow(row_count)
        
        self.table_widget.setItem(row_count, 0, QTableWidgetItem(hw_info['name']))
        
        hw_id = f"VEN_{hw_info['ven']}&DEV_{hw_info['dev']}"
        self.table_widget.setItem(row_count, 1, QTableWidgetItem(hw_id))
        
        self.table_widget.setItem(row_count, 2, QTableWidgetItem("Sürücü aranıyor..."))

    def update_driver_link(self, row, link, error):
        """İlgili satırdaki sürücü linki hücresini günceller."""
        if error:
            item = QTableWidgetItem(link) # "HATA" yazar
            item.setForeground(Qt.GlobalColor.red)
        elif link == "BULUNAMADI":
            item = QTableWidgetItem("Resmi link bulunamadı")
            item.setForeground(Qt.GlobalColor.yellow)
        else:
            item = QTableWidgetItem(link)
            # Linki mavi ve altı çizgili yaparak tıklanabilir göster
            font = item.font()
            font.setUnderline(True)
            item.setFont(font)
            item.setForeground(QGuiApplication.palette().color(QGuiApplication.Palette.ColorRole.LinkVisited))
            item.setToolTip("Linki tarayıcıda açmak için çift tıklayın.\nKopyalamak için sağ tıklayın (yakında).")
        
        self.table_widget.setItem(row, 2, item)

    def scan_finished(self):
        """Tarama bittiğinde arayüzü tekrar aktif hale getirir."""
        self.scan_button.setEnabled(True)
        self.scan_button.setText("Donanımları Tara ve Sürücüleri Bul")

    def on_cell_double_clicked(self, row, column):
        """Tablodaki bir hücreye çift tıklandığında çalışır."""
        # Sadece 3. sütuna (Link) tıklanırsa çalış
        if column != 2:
            return
            
        item = self.table_widget.item(row, column)
        if not item:
            return
        
        url = item.text()
        
        # Geçerli bir URL ise tarayıcıda aç
        if url.startswith("http://") or url.startswith("https://"):
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("Tarayıcıda Aç")
            msg_box.setText(f"Şu linki tarayıcıda açmak üzeresiniz:\n\n{url}\n\n"
                            "UYARI: Linkin resmi üretici sitesi (intel, nvidia, dell vb.) "
                            "olduğundan emin olun. Devam edilsin mi?")
            msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            msg_box.setDefaultButton(QMessageBox.StandardButton.No)
            msg_box.setIcon(QMessageBox.Icon.Warning)
            
            if msg_box.exec() == QMessageBox.StandardButton.Yes:
                try:
                    webbrowser.open_new_tab(url)
                except Exception as e:
                    self.update_status(f"Hata: Link açılamadı - {e}")

# --- Uygulamayı Başlatma ---
if __name__ == "__main__":
    if sys.platform != "win32":
        print("HATA: Bu uygulama yalnızca Windows işletim sistemlerinde çalışır.")
        sys.exit(1)
        
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
