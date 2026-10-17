# GitHub ile kalıcı kayıt ve Streamlit Cloud kurulumu

## Bir kere yapılacak ayarlar

1. GitHub'da uygulama için bir depo oluşturun. Öğretmen adları ve nöbet geçmişi depoya yazılacağı için depoyu **Private (özel)** oluşturun.
2. GitHub'da bu depoya erişebilen bir erişim anahtarı (fine-grained personal access token) üretin. Yalnızca bu depoyu seçin ve **Contents: Read and write** izni verin.
3. Uygulama dosyalarını depoya yükleyin. Streamlit giriş dosyası `nobet_uygulamasi.py`; bağımlılıklar `requirements.txt` içindedir.
4. Streamlit Community Cloud'da GitHub deponuzu, dalı (`main`) ve `nobet_uygulamasi.py` dosyasını seçerek uygulamayı yayımlayın.
5. Streamlit uygulama ayarlarında **Secrets** bölümüne aşağıdakini kendi değerlerinizle girin:

```toml
GITHUB_TOKEN = "github_pat_gercek_anahtariniz"
GITHUB_REPOSITORY = "kullanici-adiniz/depo-adi"
GITHUB_BRANCH = "main"
GITHUB_DATA_PATH = "nobet_verileri.json"
```

`GITHUB_TOKEN` değerini kaynak koda, GitHub deposuna veya herkese açık bir yere eklemeyin. `secrets.toml.example` yalnızca biçim örneğidir ve gerçek anahtar içermez.

## Nasıl çalışır?

Uygulama ilk açıldığında GitHub'daki `nobet_verileri.json` dosyasını okur. Ayarlar, öğretmenler, eski listeler ve çizelgeler kaydedildiğinde bu dosyayı günceller. Değişiklik yoksa gereksiz GitHub commit'i oluşturmaz. GitHub ayarı yoksa yerel geliştirmede eski yerel JSON dosyası kullanılmaya devam eder.

Tek kullanıcılı kullanım varsayılmıştır. GitHub deposunu özel tutun. Streamlit uygulamasına erişim de ayrıca korunmalıdır; özel GitHub deposu, yayımlanmış uygulamayı kendiliğinden özel hale getirmez.
