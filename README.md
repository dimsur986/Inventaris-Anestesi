# Inventaris Alat Anestesi — konfigurasi akses publik

## URL backend yang sudah dipasang
`https://script.google.com/macros/s/AKfycbyScrD9D_6MtxsiVygB9clXob8mtOVH0hHp2RaXDIEKWmq5JE7P8P73QWTMvHuJMrVeiA/exec`

## Cara memasang pembaruan website
1. Ekstrak ZIP ini.
2. Unggah/replace isi folder repository GitHub dengan file dari folder `Inventaris-Anestesi-main`.
3. Pastikan `index.html` berada di root repository. GitHub Pages akan menggunakan file tersebut sebagai halaman utama.
4. Commit perubahan dan tunggu GitHub Pages selesai deploy.
5. Buka website dari perangkat lain dan uji baca data. Lalu uji tambah/edit/hapus hanya dengan data percobaan yang aman.

## Penting: periksa deployment Apps Script
- Deployment harus bertipe **Web app**.
- **Execute as:** Me (pemilik script).
- **Who has access:** Anyone, sesuai pilihan akses tanpa login.
- Pastikan deployment yang di URL ini adalah deployment terbaru dan sudah memberi izin ke spreadsheet yang dituju.

## Catatan keamanan
Akses publik tanpa login berarti siapa pun yang mengetahui URL website/API dapat membaca, menambah, mengubah, dan menghapus data. Endpoint saat ini tidak memiliki autentikasi/otorisasi pengguna. Jangan gunakan untuk data rahasia atau inventaris operasional yang membutuhkan jejak audit dan pembatasan hak akses. Buat backup spreadsheet secara rutin. Untuk penggunaan rumah sakit, disarankan menambahkan autentikasi dan pembatasan hak akses sebelum penggunaan resmi.

## Catatan pengujian
URL sudah dimasukkan ke kedua file HTML dan konfigurasi tidak lagi bergantung pada `localStorage` tiap perangkat. Namun koneksi langsung ke deployment belum dapat dinyatakan berhasil hanya dari penyisipan URL; lakukan pengujian dari browser/perangkat lain. Jika muncul error, periksa akses deployment dan izin Apps Script.
