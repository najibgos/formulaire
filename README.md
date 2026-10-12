# استمارة تدقيق المعطيات الخاصة بالولي و التلميذ
# Fiche de vérification des informations du tuteur et de l'élève
### الثانوية التأهيلية علال بن عبد الله — مديرية النواصر
### Lycée qualifiant Allal Ben Abdellah — Direction provinciale de Nouaceur

**العربية:** تطبيق ويب خفيف (Streamlit) يمكّن أولياء الأمور من تدقيق معطياتهم ومعطيات أبنائهم التلاميذ. واجهة الاستمارة **بالعربية مع الكتابة من اليمين إلى اليسار**، أما **اللقب والاسم الشخصي فيُدوَّنان في خانتين: خانة بالعربية وخانة بالفرنسية** (كما في بطاقات مسار)، مع **حفظ تلقائي في ملف Excel** يمكن للإدارة معاينته وتحميله من فضاء محمي بكلمة سر.

**Français :** Application web légère (Streamlit) permettant aux tuteurs/parents de vérifier les informations les concernant ainsi que celles de leur enfant. L'interface est **entièrement en arabe (de droite à gauche)** ; le **nom et le prénom se saisissent en deux champs distincts : un champ en arabe et un champ en français** (comme sur les fiches Massar), avec **enregistrement automatique dans un fichier Excel** consultable et téléchargeable par l'administration depuis un espace protégé par mot de passe.

---

## 1. الوظائف / Fonctionnalités

| الفضاء | الوظائف |
|---|---|
| **الولي الأمر** (استمارة بالعربية) | معطيات الوالي: اللقب (بالعربية + بالفرنسية)، الاسم الشخصي (بالعربية + بالفرنسية)، رقم البطاقة الوطنية للتعريف (CIN)، رقم الهاتف، العنوان |
| | معطيات التلميذ(ة): اللقب (بالعربية + بالفرنسية)، الاسم الشخصي (بالعربية + بالفرنسية)، **تاريخ الميلاد**، رمز مسار |
| | جميع الخانات الـ **13 إلزامية** (معلمة بـ *) |
| | التحقق من صحة المعطيات: الأسماء العربية **بالحروف العربية فقط**، والأسماء الفرنسية **بالحروف اللاتينية فقط**، CIN (مثال: A123456)، الهاتف المغربي (مثال: 0612345678 أو +212...)، رمز مسار (مثال: R1300254789) |
| | منع التكرار التلقائي (نفس رمز مسار) |
| | رسالة تأكيد بعد الإرسال، ثم تفرغ الخانات لتدقيق معطيات تلميذ آخر |
| **Tuteur / parent** (interface arabe) | Tuteur : nom (champ arabe + champ français), prénom (champ arabe + champ français), n° CIN, téléphone, adresse — Élève : nom (AR + FR), prénom (AR + FR), **date de naissance**, code Massar — **13 champs obligatoires**, validation automatique (noms arabes en caractères arabes uniquement, noms français en caractères latins uniquement, formats marocains), détection des doublons par code Massar, message de confirmation puis champs vidés |
| **الإدارة** (الشريط الجانبي) | عدد التدقيقات المسجلة، معاينة جدول المعطيات، **تحميل ملف Excel** (اسم الملف يتضمن التاريخ) |
| **Administration** (barre latérale) | Compteur d'enregistrements, aperçu des données, **téléchargement du fichier Excel** (nom de fichier horodaté) |

يُنشأ الملف `data/verifications.xlsx` تلقائيا عند أول عملية تدقيق، بسطر لكل تلميذ(ة) وبـ **14 عمودا بعناوين عربية**:
تاريخ التسجيل — لقب ولي الأمر (بالعربية) — لقب ولي الأمر (بالفرنسية) — الاسم الشخصي لولي الأمر (بالعربية) — الاسم الشخصي لولي الأمر (بالفرنسية) — رقم البطاقة الوطنية للتعريف — رقم الهاتف — العنوان — لقب التلميذ(ة) (بالعربية) — لقب التلميذ(ة) (بالفرنسية) — الاسم الشخصي للتلميذ(ة) (بالعربية) — الاسم الشخصي للتلميذ(ة) (بالفرنسية) — تاريخ ميلاد التلميذ(ة) — رمز مسار.

ورقة Excel مضبوطة على العرض من اليمين إلى اليسار مع ترويسة ملوّنة مثبتة، وأعمدة الهاتف وتاريخ الميلاد ورمز مسار بتنسيق نصي للحفاظ على الأصفار في البداية.

Le fichier `data/verifications.xlsx` est créé automatiquement à la première soumission : une ligne par élève, **14 colonnes à en-têtes arabes** (nom et prénom dédoublés : version arabe + version française), feuille en affichage droite-à-gauche, en-tête coloré figé, colonnes téléphone, date de naissance et code Massar au format texte (zéros initiaux préservés).

## 2. التثبيت — Installation (مرة واحدة / une seule fois)

1. تثبيت **Python 3.9 أو أحدث** / Installer **Python 3.9 ou plus récent**: <https://www.python.org/downloads/>
   (على Windows، تحديد الخيار « Add Python to PATH »).
2. فتح نافذة الأوامر داخل هذا المجلد ثم تنفيذ / Ouvrir un terminal dans ce dossier puis exécuter :

```bash
pip install -r requirements.txt
```

## 3. تشغيل التطبيق — Lancement

- **Windows**: النقر المزدوج على `demarrer_windows.bat`
- **macOS / Linux**: `bash demarrer_mac_linux.sh`
- أو يدويا / ou manuellement :

```bash
streamlit run app.py
```

يفتح المتصفح تلقائيا على العنوان `http://localhost:8501`.
Le navigateur s'ouvre automatiquement sur `http://localhost:8501`.

## 4. فضاء الإدارة — Espace administrateur

1. فتح الشريط الجانبي « فضاء المشرف ».
2. إدخال كلمة السر الافتراضية / Saisir le mot de passe par défaut : **admin2026**.
3. الاطلاع على عدد التدقيقات والمعاينة، ثم النقر على « تحميل ملف Excel »
   (مثال اسم الملف / exemple de nom : `تدقيق_المعطيات_ثانوية_علال_بن_عبد_الله_20261013.xlsx`).

### تغيير كلمة السر — Changer le mot de passe

- عبر متغير البيئة (الطريقة الموصى بها) / via la variable d'environnement (recommandé) :
  - Windows: `set ADMIN_PASSWORD=nouveau_mot_de_passe` قبل تشغيل الملف .bat
  - macOS/Linux: `export ADMIN_PASSWORD=nouveau_mot_de_passe`
- أو بتعديل السطر `ADMIN_PASSWORD = ...` في أعلى الملف `app.py`
  / ou en modifiant la ligne `ADMIN_PASSWORD = ...` en haut du fichier `app.py`.

## 5. التعديلات السريعة — Personnalisation (ملف / fichier app.py)

- **اسم المؤسسة والمديرية**: المتغيرات `SCHOOL_NAME` و `DIRECTORATE` (السنة الدراسية تُحسب تلقائيا).
  **Établissement et direction** : variables `SCHOOL_NAME` / `DIRECTORATE` (année scolaire calculée automatiquement).
- **كلمة سر الإدارة / Mot de passe admin**: `ADMIN_PASSWORD`.
- **قواعد التحقق / Règles de validation**: `AR_NAME_RE` (حروف عربية فقط), `FR_NAME_RE` (حروف لاتينية فقط), `CIN_RE`, `PHONE_RE`, `CODE_RE`.
- **تاريخ ميلاد التلميذ / Date de naissance** : plage autorisée `min_value` / `max_value` (1980 → 2020) dans l'appel `st.date_input`.
- **عناوين الخانات / Libellés des champs**: قسم « ترويسة المؤسسة + عنوان الاستمارة ».

## 6. ملاحظات تقنية — Notes techniques

- **الهاتف المحمول / Téléphone portable** : تم ضبط العرض ليناسب الشاشات الصغيرة (النموذج متوافق مع الهواتف). الملف `.streamlit/config.toml` يخفي أزرار Streamlit التقنية (Deploy، ⋮) ويُبقي زر فتح قائمة الإدارة، التي تفتح على شكل درج منزلق على الهاتف.
  L'affichage est adapté aux petits écrans (formulaire responsive). Le fichier `.streamlit/config.toml` masque les boutons techniques de Streamlit (Deploy, ⋮) tout en conservant le contrôle qui ouvre la barre latérale, affichée en tiroir coulissant sur mobile.
- المعطيات تبقى **محلية** (لا يتم إرسال أي معلومة إلى الإنترنت).
  Les données restent **locales** (aucune information n'est envoyée sur Internet).
- إذا كان ملف Excel مفتوحا أثناء إرسال أحد الأولياء للاستمارة قد يفشل الحفظ: يُرجى إغلاق الملف بعد الاطلاع عليه.
  Si le fichier Excel est ouvert pendant une soumission, l'enregistrement peut échouer : fermer le fichier après consultation.
- لتمكين الولوج من حواسيب أخرى داخل شبكة المؤسسة / pour un accès depuis d'autres postes du réseau de l'établissement :

```bash
streamlit run app.py --server.address 0.0.0.0
```

  ثم الولوج من أي حاسوب عبر / puis accéder depuis n'importe quel poste via `http://IP-الحاسوب:8501`.
