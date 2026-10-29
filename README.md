# Pembuat Invoice – Usaha Mandiri Lift

Halaman web untuk membuat invoice (.docx) dari template Word.

## Isi folder
- `index.html` – halaman form (pilih kode invoice, isi data, unduh Word)
- `template-invoice.docx` – template Word. Tanda `{...}` di dalamnya diisi otomatis oleh halaman. Jangan ubah atau hapus tandanya.
- `vendor/` – pustaka pembuat Word (docxtemplater dan pizzip)

## Pasang di GitHub Pages
1. Buat repository baru di GitHub, lalu unggah semua isi folder ini (termasuk folder `vendor`).
2. Buka **Settings → Pages**.
3. Pada **Build and deployment**, pilih **Deploy from a branch**, branch `main`, folder `/ (root)`, lalu **Save**.
4. Tunggu 1–2 menit. Alamatnya: `https://NAMA-GITHUB.github.io/NAMA-REPO/`

## Mengubah tampilan invoice
Buka `template-invoice.docx` di Word, ubah huruf, jarak, atau logo, lalu simpan. Jangan ubah tulisan `{...}`. Unggah ulang berkas itu ke repository (ganti yang lama).

## Mengubah bawaan kode 1, 2, 3
Cari `PRESETS` di `index.html` (kode PO, tanggal PO, jatuh tempo, dan barang untuk tiap kode).

## Catatan
- Semua data diproses di browser, tidak dikirim ke server.
- Repository publik berarti siapa pun yang tahu alamatnya bisa membuka halaman dan template (termasuk nomor rekening dan data PO di dalamnya).
- Untuk dipakai di komputer sendiri tanpa internet, jalankan `python -m http.server` di folder ini lalu buka `http://localhost:8000`.
