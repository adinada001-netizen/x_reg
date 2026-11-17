=============================================================================
             9ROUTER CODEBUDDY (codebuddy-intl) AUTO-TOOLS
=============================================================================

Tools ini digunakan untuk mendaftarkan dan menghubungkan akun CodeBuddy 
ke 9Router secara otomatis (atau manual) sehingga model AI seperti:
- GLM-5.2 (cbai/glm-5.2)
- DeepSeek-V4-Pro (cbai/deepseek-v4-pro)
- Kimi-K2.7 (cbai/kimi-k2.7)
- MiniMax-M2.7 (cbai/minimax-m2.7)
dapat langsung digunakan di 9Router.

-----------------------------------------------------------------------------
CARA PENGGUNAAN CEPAT:
-----------------------------------------------------------------------------

1. INSTALLASI (Cukup dijalankan 1x):
   - Double-click '1_Setup_Install.bat'
   - Tunggu hingga dependensi dan browser otomatis selesai diunduh.

2. ISI DAFTAR AKUN GOOGLE:
   - Buka file 'Google.txt'
   - Masukkan akun Google Anda dengan format (1 per baris):
     email@gmail.com|password
     contohakun1@gmail.com|PasswordRahasia123!
     contohakun2@gmail.com|PasswordRahasia123!

3. KONFIGURASI PROXY (Opsional):
   - Buka file 'config.json'
   - Jika Anda memiliki API Key Webshare, masukkan di bagian "webshare_api_key".
   - Jika dikosongkan (""), tools akan otomatis berjalan mode Direct (tanpa proxy).

4. JALANKAN AUTO-REGISTER:
   - Double-click '2_Run_Auto_Register_Google.bat'
   - Script akan otomatis membuka browser stealth dan mendaftarkan akun.
   - Jika Google meminta verifikasi 2-digit di HP, script akan menampilkan 
     angkanya di layar terminal. Anda cukup ketuk angka tersebut di HP Anda.
   - Akun yang sukses akan langsung aktif di dashboard 9Router Anda!

5. LOGIN MANUAL / SATUAN (Jika Ingin Login Lewat X/GitHub/Apple):
   - Double-click '3_Manual_Browser_Login.bat'
   - Browser akan otomatis terbuka, login dengan metode pilihan Anda, 
     dan token akan otomatis tersambung ke 9Router.

=============================================================================
