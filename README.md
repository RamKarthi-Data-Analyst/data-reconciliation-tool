# 📊 Data Reconciliation Tool

This Excel-based VBA tool compares a main transaction list (`All Data`) against four sub-datasets (`Data 1`, `Data 2`, `Data 3`, `Data 4`) and provides reconciliation results based on **Transaction ID** and **Amount**.

---

## 🔍 Use Case

Quickly identify whether a transaction exists in sub-sheets with:
- ✅ Matching ID and Amount
- ⚠️ ID found but Amount differs
- ❌ ID not found at all

---

## 🔁 Reconciliation Logic

For each row in the **All Data** sheet:
1. If Transaction ID **and** Amount match in any subsheet →  
   ➤ `Transaction found with same amount`

2. If Transaction ID matches but Amount does not →  
   ➤ `Transaction found with different amount`

3. If Transaction ID is not found →  
   ➤ `Transaction not found`

---

## 📄 Files Included

| File | Description |
|------|-------------|
| `Data Reconciliation.xlsm` | The macro-enabled workbook with all logic implemented |
| `README.md` | This documentation file |

---

## 🛠 Tools Used
- Excel VBA (Macros)
- Sheet Looping and Range Matching
- Conditional Checks & Automation

---

## 📚 Learning Outcomes
- Built a real-world reconciliation system
- Strengthened VBA lookup logic
- Practiced condition-based reporting across multiple datasets

---

> ✅ Feel free to clone or download this for learning or adaptation to your organization’s workflow!

