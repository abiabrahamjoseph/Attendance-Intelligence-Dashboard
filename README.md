# Student Attendance Intelligence Dashboard

An institutional reporting dashboard that analyzes student attendance data directly from Excel spreadsheets (`.xlsx`, `.xls`, `.csv`). 

All processing is executed **100% client-side** inside the browser using SheetJS and Chart.js—no backend database is required, and data never leaves your computer.

---

## 🚀 Key Features

* **Drag-and-Drop Excel Upload**: Supports `.xlsx`, `.xls`, and `.csv` files.
* **Auto Column Detection**: Automatically detects columns for `Student Name`, `Attendance Status`, `Batch / Class`, `Faculty / Teacher`, and `Date`.
* **Multi-Sheet Support**: Automatically detects sheets and provides an instant sheet-switcher dropdown for multi-sheet workbooks.
* **Real-Time KPIs**: Total Records, Present, Absent, Attendance Percentage with color-coded tier progress bars (Good $\ge 90\%$, Warning $75-89\%$, Shortage $<75\%$).
* **Interactive Visual Analytics**:
  * **Attendance Distribution Donut Chart** (Present vs Absent breakdown).
  * **Batch-wise Attendance Comparison Bar Chart** (Attendance % across batches).
* **Multi-Dimensional Filters**: Filter records by Student, Batch, Status (Present / Absent / Late / Holiday), and Faculty, plus instant live search.
* **Faculty Breakdown**: Automated summary table grouped by faculty members with session counts and attendance rates.
* **Export & Print**: One-click filtered data export back to Excel (`.xlsx`) and print-optimized report styling.

---

## 🛠️ Tech Stack

* **Frontend**: HTML5, Tailwind CSS, Lucide Icons
* **Data Processing**: SheetJS (`xlsx.full.min.js`)
* **Data Visualization**: Chart.js (`chart.min.js`)
* **Local Web Server**: Node.js (`server.js`)

---

## 💻 Running Locally

### Option 1: Using Node.js (Recommended)
```powershell
node server.js
```
Then open your browser at **`http://localhost:3000/`**.

### Option 2: Direct Browser Launch
Since the dashboard is completely client-side, you can also directly open `index.html` in any modern web browser (Chrome, Edge, Firefox, Safari).

---

## 🌐 Deploying & Hosting on GitHub Pages (Free)

1. **Initialize Git and commit**:
   ```powershell
   git init
   git add .
   git commit -m "Initial commit: Student Attendance Intelligence Dashboard"
   ```

2. **Push to your GitHub repository**:
   ```powershell
   git branch -M main
   git remote add origin https://github.com/<your-username>/<your-repo-name>.git
   git push -u origin main
   ```

3. **Turn on GitHub Pages**:
   * Navigate to your repository on **GitHub.com**.
   * Go to **Settings** > **Pages** (in the left sidebar).
   * Under **Branch**, select `main` and `/ (root)`.
   * Click **Save**.

Your dashboard will be live at `https://<your-username>.github.io/<your-repo-name>/`.

---

## 🔒 Privacy & Security

* **Zero Cloud Transfer**: All spreadsheet parsing happens inside your browser's JavaScript engine via `FileReader`.
* **Institution Ready**: Complies with student data privacy requirements because no data is sent to external servers or APIs.

---

## 📄 License
MIT License.
