# GTFS_Generator_KCI_Public

Generator [GTFS](https://gtfs.org/) (General Transit Feed Specification) untuk
**KCI Commuter (Jabodetabek)**, dengan CLI yang bisa dijalankan tanpa dashboard.

Feed dibangun dari sumber jadwal resmi: workbook Excel GAPEKA KCI, atau PDF
jadwal Commuter Line yang dipublikasikan KCI.

## Isi paket

| File | Fungsi |
|------|--------|
| `generate_gtfs.py` | Entry point: parser Excel/PDF, generator, validator, dan CLI |
| `gtfs_edit.py` | Pembaca/penulis ulang `.zip` GTFS (stdlib saja) |
| `test_generate_gtfs.py` | Test suite yang mandiri, tanpa file eksternal |

## Yang dihasilkan

Feed KCI berisi 7 file:

| File | Isi |
|------|-----|
| `agency.txt` | Satu agency KCI |
| `calendar.txt` | Definisi service (`AllDay`, `Weekday`, `Weekend`) |
| `levels.txt` | Level/street (`L0` = Street) |
| `routes.txt` | 10 rute KCI |
| `stops.txt` | Stasiun yang benar-benar punya jadwal pemberhentian |
| `trips.txt` | Satu baris per perjalanan |
| `stop_times.txt` | Jam datang/berangkat per pemberhentian |

Stasiun yang selalu dilewati secara ekspres (ada di timetable, tapi tanpa jam
berhenti) tidak ditulis ke `stops.txt`, karena tidak ada pemberhentian terjadwal
di sana. Pada workbook GAPEKA terbaru, `JIS` dan `KAT` termasuk kategori ini.

## Kebutuhan

| Sumber | Kebutuhan |
|--------|------------|
| Mode demo (default) | Tidak ada, Python 3.9+ |
| `--excel` | `openpyxl` |
| `--pdf` / `--pdf-url` | `pymupdf` |

```bash
pip install openpyxl    # untuk --excel
pip install pymupdf     # untuk --pdf / --pdf-url
```

## Cara pakai

### Feed KCI dari PDF resmi

```bash
python generate_gtfs.py --pdf-url --out output --zip
```

### Feed KCI dari Excel

```bash
python generate_gtfs.py --excel "A12_B. Jadwal JABO GAPEKA 2025.xlsx" --out output --zip
```

### Filter hari

```bash
python generate_gtfs.py --pdf-url --service Weekday --service Weekend --out output --zip
```

### Mode demo (tanpa file apa pun)

```bash
python generate_gtfs.py --out output --zip
```

### Opsi

| Flag | Default | Keterangan |
|------|---------|------------|
| `--excel` | - | Workbook GAPEKA KCI (`.xlsx`) |
| `--pdf` | - | PDF jadwal KCI lokal |
| `--pdf-url` | - | URL PDF jadwal; default: PDF Jabodetabek terkini |
| `--service` | semua | Batasi ke `AllDay`, `Weekday`, atau `Weekend`; bisa diulang |
| `--out` | `output` | Folder keluaran atau path `.zip` |
| `--zip` | `False` | Tulis sekaligus arsip `.zip` |
| `--routes` | `3` | Demo: jumlah rute |
| `--stops-per-route` | `6` | Demo: pemberhentian per rute |
| `--trips-per-day` | `5` | Demo: perjalanan per hari per rute |
| `--seed` | `42` | Demo: seed agar reproducible |
| `--start-date` | auto | Demo: tanggal mulai feed (`YYYY-MM-DD`) |
| `--config` | - | Demo: file konfigurasi JSON |

Pilih **salah satu** dari `--excel`, `--pdf`, atau `--pdf-url`. Tanpa salah
satunya, tool berjalan dalam mode demo.

## Manajemen stasiun

Feed yang sudah jadi bisa diinspeksi dan diubah tanpa membuat ulang jadwal.

```bash
# Daftar stasiun pada database
python generate_gtfs.py stations

# Isi feed
python generate_gtfs.py show-stops --feed output.zip

# Stop di stops.txt yang tidak pernah dipakai stop_times
python generate_gtfs.py detect-orphans --feed output.zip

# Cek feed sebelum dipakai sebagai master
python generate_gtfs.py classify --feed output.zip

# Nonaktifkan stasiun
python generate_gtfs.py disable-stops --feed output.zip --stop TPK --out output_tanpa_tpk.zip --zip

# Aktifkan kembali, memulihkan stop_times dari master
python generate_gtfs.py activate-stops --feed output_tanpa_tpk.zip \
    --master output.zip --out output_lengkap.zip --zip
```

| Flag | Default | Keterangan |
|------|---------|------------|
| `--feed` | wajib | Feed `.zip` masukan |
| `--stop` | - | `stop_id` yang akan diproses; bisa diulang. Kosongkan untuk memakai daftar tersimpan |
| `--out` | `output.zip` | Keluaran `.zip` |
| `--zip` | `False` | Tulis `.zip` |
| `--state` | `inactive_stops.json` | JSON yang menyimpan daftar stasiun nonaktif |
| `--master` | `gtfs_master.zip` | Master untuk memulihkan `stop_times` saat reaktivasi |
| `--backfill` | `False` | `disable-stops` saja: pulihkan `stop_times` dari master |
| `--active-only` | `False` | `stations` saja: sembunyikan stasiun nonaktif |

`classify` berguna sebelum menimpa master. `rebuild_gtfs_without_inactive`
menyisakan `stop_times` hanya untuk stasiun aktif, sehingga upload setengah jadi
bisa lolos dari pemeriksaan orphan. `classify` menutup celah itu dengan
menilai apakah feed masih memuat mayoritas stasiun yang diharapkan.

## Test

```bash
python -m unittest test_generate_gtfs -v
```

Seluruh fixture dibuat di memori, jadi suite berjalan tanpa file jadwal dan
tanpa koneksi internet.

## Data stasiun

Nama stasiun dan koordinat berasal dari data KCI dan sengaja dipublikasikan
agar feed yang dihasilkan bisa langsung dipakai. `stop_code` yang dipakai
(`10001`-`10085`) adalah nomor sintetis, bukan kode resmi KCI.

## Proyek lain

Pasangkan dengan
[GTFS_Monitoring_Public](https://github.com/rahardikanugroho/GTFS_Monitoring_Public)
untuk memvalidasi feed yang dihasilkan.

## License

MIT
