# Gümrük Beyanname Sorgulama Otomasyonu

Bu proje, Excel listesindeki Gümrük Beyanname (GCB) ve ETGB numaralarını gümrük sorgulama sistemine göndererek, **İntaç Tarihi** veya **Kapanmamış** durum bilgilerini alan ve Excel'i gerçek zamanlı olarak güncelleyen bir web otomasyon aracıdır.

## 🚀 Özellikler
*   **Tamamen Headless (Saf HTTP)**: Tarayıcı (Chrome/Selenium) gerektirmez, çok hızlıdır ve sunucu kaynaklarını tüketmez.
*   **Gelişmiş CAPTCHA Çözücü**: Matematiksel CAPTCHA görsellerini piksel analiziyle %95+ doğrulukla çözer.
*   **Paralel Sorgulama (Multi-threading)**: Gümrük limitlerine takılmadan birden fazla iş parçacığıyla sorgu yapar.
*   **Otomatik Cooldown (Sıra/Bekleme) Yönetimi**: Rate-limit durumlarında (5 dakika) otomatik bekler ve geri sayım yapar.
*   **Dinamik Excel Desteği**: Sütunları otomatik algılar ve günceller.
*   **Canlı Terminal Akışı**: Detaylı loglama ve işlem takibi.

---

## 🛠️ Yerel Kurulum ve Çalıştırma

### 1. Gereksinimler
Sisteminizde Python 3.10 veya üzeri sürümün kurulu olduğundan emin olun.

### 2. Kurulum
Terminali proje dizininde açın ve bağımlılıkları yükleyin:
```bash
pip install -r requirements.txt
```

### 3. Çalıştırma
Uygulamayı başlatın:
```bash
python main.py
```
Sunucu başladığında tarayıcınızda **http://127.0.0.1:8000** adresini açarak uygulamayı kullanabilirsiniz.

---

## 📦 GitHub'a Yükleme Rehberi

Projenizi kendi GitHub hesabınıza yüklemek için aşağıdaki adımları sırayla takip edin:

1.  **Git'i Başlatın**:
    Proje klasörünüzde terminali açın:
    ```bash
    git init
    git branch -M main
    ```

2.  **Dosyaları Ekleyin ve Commit Edin**:
    ```bash
    git add .
    git commit -m "İlk commit: Gümrük sorgulama otomasyonu hazır"
    ```

3.  **GitHub'da Yeni Bir Depo (Repository) Oluşturun**:
    *   [github.com](https://github.com) adresine gidin ve giriş yapın.
    *   Sağ üstteki **+** butonuna tıklayıp **New repository** seçeneğini seçin.
    *   Depoya bir isim verin (Örn: `gumruk-otomasyon`) ve **Create repository** butonuna tıklayın.

4.  **Uzak Depoyu Ekleyin ve Dosyaları Gönderin**:
    Oluşturduğunuz deponun sayfasındaki Git adresini kopyalayıp terminale yapıştırın:
    ```bash
    git remote add origin https://github.com/KULLANICI_ADINIZ/gumruk-otomasyon.git
    git push -u origin main
    ```

---

## ☁️ Sunucuda Canlıya Alma (Deployment)

Projenizi GitHub'a yükledikten sonra, uygulamayı internet üzerinden 7/24 çalışacak şekilde canlıya alabilirsiniz.

### Seçenek 1: Render.com (Önerilen & Ücretsiz Seçenek)
1.  [Render.com](https://render.com) adresine ücretsiz üye olun.
2.  Dashboard'dan **New +** > **Web Service** seçeneğine tıklayın.
3.  GitHub hesabınızı bağlayın ve oluşturduğunuz `gumruk-otomasyon` reposunu seçin.
4.  Aşağıdaki ayarları yapın:
    *   **Runtime**: `Python`
    *   **Build Command**: `pip install -r requirements.txt`
    *   **Start Command**: `uvicorn main:app --host 0.0.0.0 --port $PORT`
5.  **Deploy Web Service** butonuna tıklayın. Render size `https://gumruk-otomasyon.onrender.com` gibi canlı bir web linki verecektir.

### Seçenek 2: Railway.app
1.  [Railway.app](https://railway.app) adresine üye olun.
2.  **New Project** > **Deploy from GitHub repository** diyerek reponuzu seçin.
3.  Railway projeyi otomatik algılayıp doğrudan canlıya alacaktır.
