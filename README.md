# 🧠 VBA_PARSER Solution

**VBA_PARSER** este o suită de instrumente pentru analiză, export și vizualizare a codului VBA din fișiere Microsoft Access (`.mdb` / `.accdb`).

Proiectul este scris integral în **VB.NET (Visual Basic)** și organizat pe 3 componente majore:

---

## 🧩 Structură soluție

```
VBA_PARSER/
│
├── 📁 VBA_CORE/          → Motorul principal de analiză și tipuri comune
│     ├── modTypes.vb
│     ├── Logger.vb
│     └── (alte module)
│
├── 📁 VBA_TOKENIZER/     → Parserul și analizorul de cod VBA
│     ├── Program.vb
│     ├── modExporter.vb
│     ├── modFileParser.vb
│     ├── modReferenceResolver.vb
│     └── (alte module logice)
│
├── 📁 VBA_VIEWER/        → Interfața WPF (UI) pentru explorarea structurii proiectului
│     ├── MainWindow.xaml
│     ├── MainWindow.xaml.vb
│     └── App.xaml
│
└── 📄 VBA_PARSER.sln     → Soluția Visual Studio principală
```

---

## ⚙️ Cerințe

- **Visual Studio 2022** (sau mai nou)
- **.NET Framework 4.8**
- **Microsoft Access** instalat (pentru export via `Microsoft.Office.Interop.Access`)
- Pachetul **EPPlus** (pentru export Excel)
- Permisiuni de acces la fișiere `.accdb` / `.mdb`

---

## 🚀 Componentele soluției

### **🔹 VBA_CORE**
- Definirea tipurilor de date partajate (`modTypes.vb`)
- Sistemul de logging (`Logger.vb`)
- Structuri container (`ModuleContainer`, `FormReportContainer`)
- Utilitare comune pentru toate proiectele

---

### **🔹 VBA_TOKENIZER**
- Exportă module, formulare și rapoarte din Access
- Parsează cod VBA și construiește o bază de simboluri
- Identifică metode, variabile, constante și relații
- Generează rapoarte (`FunctionsIndex.json`, `.html`, `.xlsx`)
- Salvează și încarcă cache-ul de analiză (`cache.dat`)

---

### **🔹 VBA_VIEWER**
- Interfață WPF care afișează structura completă a proiectului VBA:
  - Module, clase, formulare și rapoarte
  - Funcții, variabile, constante, tokeni, linii de cod
  - Lazy-loading pentru performanță
- Posibilitate de căutare rapidă în metode
- Detalii contextuale pentru fiecare nod

---

## 💻 Cum se rulează

1. Deschide soluția `VBA_PARSER.sln` în Visual Studio.
2. Configurează proiectul de start: **VBA_TOKENIZER** (sau **VBA_VIEWER** pentru UI directă).
3. Asigură-te că fișierul Access (`.accdb` sau `.mdb`) e configurat în `DB_PATH`.
4. Rulează (F5).

---

## 🧠 Funcționalități cheie

| Componentă       | Descriere principală |
|------------------|----------------------|
| 🔹 Export Access | Extrage cod VBA din obiecte Access (Modules, Forms, Reports) |
| 🔹 Tokenizare | Desparte codul în tokeni și construiește modelul sintactic |
| 🔹 Analiză simboluri | Detectează funcții, variabile, constante, tipuri, proprietăți |
| 🔹 Referințe | Urmărește dependențele și apelurile între module |
| 🔹 Rapoarte | Generează HTML, JSON și Excel cu rezultatele |
| 🔹 Vizualizare WPF | Explorează codul exportat într-un TreeView interactiv |

---

## 🧰 Comenzi CLI

Poți rula aplicația principală (`VBA_TOKENIZER.exe`) cu parametri:

| Argument | Descriere |
|-----------|------------|
| `--cache` | Încarcă analiza din cache dacă există |
| `--nothreads` | Dezactivează procesarea paralelă |
| `--verbose` | Afișează log detaliat |
| `--quiet` | Log minimal |
| `--noopen` | Rulează fără a deschide UI-ul WPF |

---

## 📊 Output generat

După rulare, în directorul `ExportVBA` se vor genera fișiere:
- `FunctionsIndex.json` / `.txt` / `.html`
- `Functions_AI.json` / `Code_AI.txt`
- `VBA_Structure.xlsx`
- `cache.dat`

---

## 👩‍💻 Autor

**Adelina Trandafir**  
📍 România  
💬 GitHub: [github.com/adelinatrandafir](https://github.com/adelinatrandafir)

---

## 🪄 Licență

Acest proiect este oferit sub licența **MIT**, permițând utilizarea și modificarea liberă, cu menționarea autorului original.
