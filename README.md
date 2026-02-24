# 🧾 Meesho Reconciliation Tool

Automated price reconciliation tool for Meesho seller accounts (YG · PE · AG), built with Python & Streamlit.

## 🚀 Live App
> Deploy on [Streamlit Cloud](https://streamlit.io/cloud) for free.

---

## 📋 What It Does

1. **Uploads** Meesho order CSVs (YG, PE, AG accounts)
2. **Translates** Meesho SKUs → OMS SKUs using the Replace SKU mapping
3. **Checks** if SKU is in the Closed SKU list → uses the lesser price (Sheet1 vs Sheet2) as PWN+10%
4. **Looks up** PWN+10% price from the PWN reference file (flags "SKU Not Found" if missing)
5. **Calculates** per order:
   - `Supplier Discounted Price × Qty`
   - `Update PWN+10% × Qty`
   - `Difference Amount` = above two subtracted (Profit ✅ or Loss ❌)
6. **Exports** a colour-coded Excel report with separate sheets per account + a Summary sheet

---

## 📁 File Structure

```
app.py               ← Main Streamlit app
requirements.txt     ← Python dependencies
README.md            ← This file
```

---

## 🖥️ Run Locally

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_USERNAME/meesho-reconciliation.git
cd meesho-reconciliation

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the app
streamlit run app.py
```

---

## ☁️ Deploy on Streamlit Cloud (Free)

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **New App** → select your repo → set `app.py` as the main file
4. Click **Deploy** ✅

---

## 📂 Reference Files Required (upload in sidebar)

| File | Purpose |
|------|---------|
| `Replace_SKU.xlsx` | Maps Meesho SKUs → OMS SKUs (3 sheets: YG, Pushpa, AG) |
| `PWN_10_.xlsx` | PWN+10% expected price per OMS Child SKU |
| `meesho_closed_sku.xlsx` | Closed/discontinued SKUs with prices (Sheet1 & Sheet2) |

## 📂 Order CSV Files (upload in main area)

- Filenames must contain `_YG`, `_PE`, or `_AG` for auto account detection
- Example: `Orders_2026-02-21_YG.csv`

---

## 🎨 Excel Output Colour Guide

| Colour | Meaning |
|--------|---------|
| 🟢 Green | Profit (Difference ≥ 0) |
| 🔴 Red | Loss (Difference < 0) |
| 🔵 Blue | Closed SKU row |
| 🟡 Yellow | SKU Not Found in PWN file |

---

## 🔄 Reconciliation Logic

```
For each order row:
  1. Build OMS SKU = Replace_SKU lookup(Meesho SKU + Size)
  2. If OMS SKU in Closed SKU list:
       PWN+10% = min(Sheet1 price, Sheet2 price) for that SKU
       SKU Status = "Closed"
  3. Else:
       PWN+10% = lookup from PWN_10_ file
       If not found → "SKU Not Found"
       SKU Status = "On Going"
  4. Difference = (Supplier Discounted Price × Qty) - (PWN+10% × Qty)
```
