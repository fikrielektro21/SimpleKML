# SimpleKML
# Ini untuk membuat Peta KML dengan Python menggunakan Gui untuk enntry Excel 

 # di bagian ini untuk membuat deskripsi yang di ambil dari masing-masing kolom dan di tampilkan pada Deskripsi Lokasi BTS

def get_site_info_by_id(bts_coords, site_id):
    """Fungsi untuk mengambil informasi berdasarkan SITE ID."""
    site_info = bts_coords.get(site_id.strip().upper())
    if site_info:
        return (
            f"<b>SITE ID :</b> {site_id}<br>"
            f"<b>SITE NAME :</b> {site_info['site_name']}<br>"
            f"<b>LINK ROUTE 2G :</b> {site_info['link_transport_2g']}<br>"
            f"<b>LINK ROUTE 4G :</b> {site_info['link_transport_4g']}<br>"
            f"<b>Remark Transport 2G :</b> {site_info.get('transport2g', 'Tidak tersedia')}<br>"
            f"<b>Remark Transport 4G :</b> {site_info.get('transport4g', 'Tidak tersedia')}<br>"
            f"<b>IP 2G :</b> {site_info.get('ip_2g', 'Tidak tersedia')}<br>"
            f"<b>IP 4G :</b> {site_info.get('ip_4g', 'Tidak tersedia')}<br>"
            f"<b>IP OAM :</b> {site_info.get('ip_oam', 'Tidak tersedia')}<br>"
            f"<b>BAND :</b> {site_info['band']}<br>"
        )
    return "<b>Data tidak ditemukan</b><br>"


# DI BAGIAN INI UNTUK MEMBUAT MARKER UNTUK SETIAP TITIK BTS, AGAR LEBIH MUDAH DIBEDAKAN BERDASARKAN JENIS TRANSPORT YANG DIGUNAKAN TRANSPORTNYA HARUS MEMILIKI KOLOM TERSENDIRI DAN SEBARIS DENGAN SITE ID, DAN INI DISEBUT SEBAGAI ATRIBUT SITEID

def get_marker_icon(marker_type, transport):
    """Fungsi untuk mendapatkan URL ikon marker berdasarkan jenis marker dan transport."""
    if transport == "ONT":
        return "https://maps.google.com/mapfiles/kml/pushpin/red-pushpin.png"
    elif transport == "NEC":
        return "https://maps.google.com/mapfiles/kml/pushpin/ltblu-pushpin.png"
    elif transport == "L2SW":
        return "https://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png"
    elif transport == "OFDM":
        return "https://maps.google.com/mapfiles/kml/pushpin/grn-pushpin.png"
    elif transport == "ROUTER":
        return "https://maps.google.com/mapfiles/kml/pushpin/purple-pushpin.png"
    else:
        return "https://maps.google.com/mapfiles/kml/pushpin/wht-pushpin.png"


# INI FUNGSI UNTUK MEMBUAT KML SYNTAX DARI PAKAGE SimpleKML


def create_kml_from_data(bts_file, output_kml):
    # Load data BTS dari Excel
    bts_df = pd.read_excel(bts_file)

    # Buat dictionary untuk koordinat berdasarkan SITE ID
    bts_coords = {
        row["SITE ID"].strip().upper(): {
            "coords": (row["Longitude"], row["Latitude"]),
            "site_name": row["SITE NAME"],
            "link_transport_2g": row.get("LINK TRANSPORT 2G", "Tidak tersedia"),
            "link_transport_4g": row.get("LINK TRANSPORT 4G", "Tidak tersedia"),
            "transport2g": row.get("Remark Transport 2G", "Tidak tersedia"),
            "transport4g": row.get("Remark Transport 4G", "Tidak tersedia"),
            "ip_2g": row.get("IP 2G", "Tidak tersedia"),
            "ip_4g": row.get("IP 4G_UP", "Tidak tersedia"),
            "ip_oam": row.get("IP OAM", "Tidak tersedia"),
            "band": row["BAND"],
            "marker_type": row.get("TRANSPORT", ""),
            "full_link_route_4g": row.get("LINK ROUTE 4G", "Tidak tersedia"),
        }
        for _, row in bts_df.iterrows()
    }

    # Buat objek KML
    kml = simplekml.Kml()
# dengan ini sebenarnya kml sudah di sajikan tetapi akan membuat secara Defaultnya 

# Tambahkan marker untuk setiap BTS marker yang di ambil di function di atas akan di proses di sini
    for site_id, info in bts_coords.items():
        marker_name = f"{site_id}, {info['site_name']}\n"
        point_kml = kml.newpoint(name=marker_name, coords=[info["coords"]])

  # Memanggil fungsi untuk mendapatkan deskripsi berdasarkan SITE ID
        description = get_site_info_by_id(bts_coords, site_id)
        point_kml.description = description

        # Mendapatkan transportasi untuk menentukan marker
        marker_type = str(info["marker_type"])
        transport = marker_type

        # Menggunakan fungsi untuk menentukan ikon marker berdasarkan jenis transportasi
        icon_href = get_marker_icon(marker_type, transport)
        point_kml.style.iconstyle.icon.href = icon_href

    # Proses Link Route 4G
    for _, row in bts_df.iterrows():
        if "LINK ROUTE 4G" not in row or pd.isna(row["LINK ROUTE 4G"]):
            continue  # Lewati jika kolom kosong

        route = row["LINK ROUTE 4G"]
        points = route.split('<NEC>')  # Pisahkan berdasarkan <NEC>
        coords = []

        for point in points:
            # Ekstrak SITE ID sebelum '_'
            site_id = point.split('_')[0].strip().upper()

            if site_id in bts_coords:
                coords.append(bts_coords[site_id]["coords"])
    
                

        # Tambahkan garis rute jika memiliki koordinat
        if len(coords) > 1:
            line = kml.newlinestring(
                name=f"Route {route}",
                coords=coords
            )
            line.style.linestyle.color = (
                simplekml.Color.green if "<NEC>" in route else simplekml.Color.blue
            )
            line.style.linestyle.width = 2

    # Simpan file KML
    kml.save(output_kml)
    print(f"File KML berhasil disimpan ke: {output_kml}")
  

# File input dan output
bts_file = 'D:/ganti dengan nama file excel anda.xlsx'
output_kml = 'D:/ganti dengan nama file kml yang anda inginkan.kml'


# Panggil fungsi untuk membuat KML
create_kml_from_data(bts_file, output_kml)

# selebihnya yang bagian bawah ini untuk membuat gui sederhana, jangan lupa untuk menyesuaikan kolom yang anda miliki dengan yang akan di baca oleh PY, dan gunakan file lain untuk percobaan agar data yang anda miliki tidak hilang apabila ada kekeliruan

class DataConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Konverter Excel ke KML")
        self.setGeometry(100, 100, 800, 600)

        # Layout utama
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QHBoxLayout()  # Gunakan QHBoxLayout
        self.central_widget.setLayout(self.main_layout)

        # Pembagi (Splitter)
        self.splitter = QSplitter(Qt.Horizontal)
        self.main_layout.addWidget(self.splitter)

        # Layout untuk input
        self.input_widget = QWidget()
        self.input_layout = QVBoxLayout(self.input_widget)
        self.splitter.addWidget(self.input_widget)

        # Layout untuk tabel
        self.table_widget = QWidget()
        self.table_layout = QVBoxLayout(self.table_widget)
        self.splitter.addWidget(self.table_widget)

        # Jumlah input fields yang baru
        self.input_fields = {}

        # Pemilihan file
        self.file_label = QLabel("Tidak ada file yang dipilih")
        self.input_layout.addWidget(self.file_label)
        self.select_file_button = QPushButton("Pilih File Excel")
        self.select_file_button.clicked.connect(self.load_file)
        self.input_layout.addWidget(self.select_file_button)

        # Kontainer untuk input dinamis
        self.dynamic_input_layout = QFormLayout()
        self.input_layout.addLayout(self.dynamic_input_layout)

        # Tombol untuk menambahkan data baru
        self.add_data_button = QPushButton("Tambah Data Baru")
        self.add_data_button.clicked.connect(self.add_data)
        self.input_layout.addWidget(self.add_data_button)

        # Tombol aksi
        self.save_button = QPushButton("Simpan Data")
        self.save_button.clicked.connect(self.save_data)
        self.input_layout.addWidget(self.save_button)

        self.generate_kml_button = QPushButton("Hasilkan File KML")
        self.generate_kml_button.clicked.connect(self.generate_kml)
        self.input_layout.addWidget(self.generate_kml_button)

        # Tampilan tabel
        self.data_table = QTableWidget()
        self.table_layout.addWidget(self.data_table)

        # Status internal
        self.dataframe = None

    def load_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Pilih File Excel", "", "Excel Files (*.xls *.xlsx)", options=options)
        if file_path:
            self.file_label.setText(file_path)
            try:
                self.dataframe = pd.read_excel(file_path)
                self.display_data()
                self.create_input_fields()  # Panggil fungsi untuk membuat input fields
            except Exception as e:
                QMessageBox.critical(self, "Kesalahan", f"Gagal memuat file: {str(e)}")

    def display_data(self):
        if self.dataframe is not None:
            self.data_table.setRowCount(0)
            self.data_table.setColumnCount(len(self.dataframe.columns))
            self.data_table.setHorizontalHeaderLabels(self.dataframe.columns)

            for row_idx, row in self.dataframe.iterrows():
                self.data_table.insertRow(row_idx)
                for col_idx, value in enumerate(row):
                    self.data_table.setItem(row_idx, col_idx, QTableWidgetItem(str(value)))

    def create_input_fields(self):
        # Bersihkan input sebelumnya
        for input_field in self.input_fields.values():
            self.dynamic_input_layout.removeWidget(input_field)
            input_field.deleteLater()  # Menghapus widget dari memori

        # Buat line edit untuk setiap kolom dalam DataFrame
        self.input_fields.clear()  # Reset input_fields pada setiap pemanggilan
        for col in self.dataframe.columns:
            input_field = QLineEdit()
            input_field.setPlaceholderText(f"Masukkan {col}")
            self.input_fields[col] = input_field
            self.dynamic_input_layout.addRow(QLabel(col), input_field)  # Menggunakan QFormLayout untuk label dan input

    def add_data(self):
        if self.dataframe is not None:
            try:
                new_data = {}
                # Ambil nilai dari input field untuk setiap kolom
                for col in self.dataframe.columns:
                    input_field = self.input_fields[col]
                    new_data[col] = input_field.text().strip()
                    Value = input_field.text().strip()

                    # Jika kolom adalah Longitude atau Latitude, pastikan itu adalah float
                    if col in ["Longitude", "Latitude"]:
                        if Value:
                            new_data[col] = float(new_data[col])
                        else:
                            raise ValueError(f"{col} tidak boleh kosong")
                    
                    else:
                        new_data[col] = Value

                # Tambahkan data baru pada DataFrame
                
                new_row = pd.DataFrame([new_data])
                self.dataframe = pd.concat([self.dataframe, new_row], ignore_index=True)
                self.display_data()
                QMessageBox.information(self, "Sukses", "Data baru berhasil ditambahkan!")
            except ValueError as e:
                QMessageBox.critical(self, "Kesalahan", f"kesalahan pada input : {str(e)}")
            
            except ValueError as e:
                QMessageBox.critical(self, "Kesalahan", f"Pastikan Longitude dan Latitude adalah angka yang valid.: {str(e)}")
        else:
            QMessageBox.warning(self, "Peringatan", "Silakan muat file Excel terlebih dahulu.")

    def save_data(self):
        if self.dataframe is not None:
            options = QFileDialog.Options()
            save_path, _ = QFileDialog.getSaveFileName(self, "Simpan Data", "", "Excel Files (*.xls *.xlsx)", options=options)
            if save_path:
                try:
                    self.dataframe.to_excel(save_path, index=False)
                    QMessageBox.information(self, "Sukses", "Data berhasil disimpan!")
                except Exception as e:
                    QMessageBox.critical(self, "Kesalahan", f"Gagal menyimpan file: {str(e)}")

    def generate_kml(self):
        if self.dataframe is not None:
            options = QFileDialog.Options()
            kml_path, _ = QFileDialog.getSaveFileName(self, "Simpan File KML", "", "KML Files (*.kml)", options=options)
            
            if kml_path:
                try:
                    create_kml_from_data(self.file_label.text(), kml_path)
                    QMessageBox.information(self, "Sukses", f"File KML berhasil dihasilkan di {kml_path}!")
                except Exception as e:
                    QMessageBox.critical(self, "Kesalahan", f"Gagal menghasilkan file KML: {str(e)}")
        else:
            QMessageBox.warning(self, "Peringatan", "Silakan muat file Excel terlebih dahulu.")

# Titik masuk utama
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = DataConverterApp()
    main_window.show()
    sys.exit(app.exec_())
