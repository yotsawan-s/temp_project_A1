# 📂 Part Reconcile — เอกสารข้อมูล (Data Catalog)

> เอกสารสรุปข้อมูลโครงการ **BOT Reconciliation** จากไฟล์ `MockUp_File.xlsx`
> (ชีท `0_Structure`, `1_Flow_Map`, `3_DataSource`) — แบ่งเป็นหมวดหมู่เพื่อใช้อ้างอิงในการออกแบบเครื่องมือ
>
> **วัตถุประสงค์:** ตรวจสอบว่าข้อมูล Transaction ใน **A1_System** บันทึกตรงกับ **Product Database** ครบถ้วน-ถูกต้อง ตาม *Data Transaction Date* ที่ผู้ใช้ระบุ ก่อนส่งต่อหน่วยงานอื่น

---

## สารบัญ

1. [ภาพรวมการ Reconcile](#1-ภาพรวมการ-reconcile)
2. [หมวด A — โครงสร้างไฟล์ (File Structure)](#หมวด-a--โครงสร้างไฟล์-file-structure)
3. [หมวด B — Product & Category (Flow Map)](#หมวด-b--product--category-flow-map)
4. [หมวด C — Data Source (Field ที่ใช้)](#หมวด-c--data-source-field-ที่ใช้)
5. [หมวด D — Reconcile Rules (กฎการจับคู่)](#หมวด-d--reconcile-rules-กฎการจับคู่)
6. [หมวด E — มาตรฐานคำศัพท์ & DSL](#หมวด-e--มาตรฐานคำศัพท์--dsl)

---

## 1. ภาพรวมการ Reconcile

ตรวจ **2 มุมมอง** ต่อ 1 Match Key:

| มุม | ทิศทาง | สิ่งที่จับ | ความหมาย |
|----|--------|-----------|----------|
| **มุม 1** | A1_System → Database | **EXTRA in A1** | มีใน A1 แต่ไม่มีใน Database (บันทึกเกิน) |
| **มุม 2** | Database → A1_System | **MISSING in A1** | มีใน Database แต่ไม่มีใน A1 (บันทึกขาด) |
| **ค่า** | Key ตรงทั้งคู่ | **VALUE MISMATCH** | เทียบ CCY + Amount + Date แล้วไม่ตรง |

- จับคู่ **รายบรรทัด** ด้วย Match Key — **ไม่รวมยอด (no SUM)**
- ผลลัพธ์ที่เป็น Error → ผูกกับ Key เป็น **ERROR Message** นำกลับไปแก้ที่ A1_System (= ไฟล์ `1_Bxxx`)

---

## หมวด A — โครงสร้างไฟล์ (File Structure)

> ที่มา: ชีท `0_Structure` | Root: `1_input/`

### A.1 กลุ่มไฟล์ทั้งหมด (รวม ~29 ไฟล์)

| กลุ่ม (Folder) | ไฟล์ | จำนวน | บทบาท |
|---------------|------|:-----:|-------|
| `AF1_System` | `1_A001` – `1_A009` | 9 | **A1_System (main)** — ระบบที่ตรวจ (เกิน/ขาด) |
| `AF1_System` | `1_B001` – `1_B009` | 9 | **A1_System (ERROR output)** — ไฟล์ผลลัพธ์ `[Key, ERROR Message]` |
| `D_Database` | `2_D001` – `2_D005` | 5 | **ฐานข้อมูล Product D** |
| `O_Database` | `3_O001` – `3_O003` | 3 | **ฐานข้อมูล Product O** |
| `C_Database` | `4_C001` – `4_C003` | 3 | **ฐานข้อมูล Product C** |

### A.2 จุดสำคัญ
- ไฟล์ **`1_Bxxx`** มีคอลัมน์ Select เพียง `[Key, ERROR Message]` → **คือไฟล์ปลายทางที่เก็บผลการ Reconcile** เพื่อนำกลับไปแก้ที่ A1_System
- Prefix ตัวเลขบอกกลุ่ม: `1_` = A1_System, `2_` = DB-D, `3_` = DB-O, `4_` = DB-C

---

## หมวด B — Product & Category (Flow Map)

> ที่มา: ชีท `1_Flow_Map` | รวม **3 Products, 21 Sets** | แต่ละ Set = สายโซ่ `Source#1 → Source#4` พร้อม `Trans.Map` (cardinality)

### B.1 Product D (12 Sets)

#### Category D1 — base = `2_D001`
| Set | สายโซ่ (Chain) | Trans.Map |
|:--:|----------------|-----------|
| 1 | `2_D001` → `1_A009` → `1_B009` | 1 : 1 : n |
| 2 | `2_D001` → `1_A005` → `1_B002` | 1 : 1 : n |
| 3 | `2_D001` → `1_A001` → `1_B003` | 1 : 1,2 : n |
| 4 | `2_D001` → `1_A002` → `1_B004` | 1 : 1 : n |

#### Category D2 — base = `2_D002`
| Set | สายโซ่ (Chain) | Trans.Map |
|:--:|----------------|-----------|
| 5 | `2_D002` → `1_A003` → `2_D001` → `1_B006` | 1 : 1 : 1 : n |

#### Category D3 — base = `2_D003`
| Set | สายโซ่ (Chain) | Trans.Map |
|:--:|----------------|-----------|
| 6 | `2_D003` → `1_A001` → `1_B003` | 1 : 1 : n |
| 7 | `2_D003` → `1_A002` → `1_B004` | 1 : 1 : n |
| 8 | `2_D003` → `1_A004` → `1_B001` | 1 : 1 : n |
| 9 | `2_D003` → `1_A005` → `1_B002` | 1 : 1 : n |
| 10 | `2_D003` → `1_A008` → `1_B008` | 1 : 1 : n |
| 11 | `2_D003` → `1_A006` → `1_B005` | 1 : n : n |
| 12 | `2_D003` → `2_D004` → `2_D005` → `1_A005` | n : 1 : 1 : ? |

### B.2 Product O (6 Sets)

#### Category O1 — base = `3_O001`
| Set | สายโซ่ (Chain) | Trans.Map |
|:--:|----------------|-----------|
| 1 | `3_O001` → `1_A007` → `1_B007` | 1 : 1 : n |
| 2 | `3_O001` → `1_A005` → `1_B002` | 1 : 1 : n |
| 3 | `3_O001` → `1_A001` → `3_O003` → `1_B003` | 1 : 1 : 1 : n |
| 4 | `3_O001` → `1_A002` → `1_B004` | 1 : 1 : n |

#### Category O2 — base = `3_O002`
| Set | สายโซ่ (Chain) | Trans.Map |
|:--:|----------------|-----------|
| 5 | `3_O002` → `1_A005` → `1_B002` | 1 : 2 : n |

#### Category O3 — base = `3_O003`
| Set | สายโซ่ (Chain) | Trans.Map |
|:--:|----------------|-----------|
| 6 | `3_O003` → `1_A008` → `1_B008` | 1 : 1 : n |

### B.3 Product C (3 Sets)

#### Category C1 — base = `1_A008`
| Set | สายโซ่ (Chain) | Trans.Map |
|:--:|----------------|-----------|
| 1 | `1_A008` → `4_C001` | 1 : 2 |
| 2 | `1_A008` → `4_C002` → `4_C003` | 1 : 1 : 1 |
| 3 | `1_A008` → `1_B008` | 1 : n |

---

## หมวด C — Data Source (Field ที่ใช้)

> ที่มา: ชีท `3_DataSource` | แสดงเฉพาะ Flag = **Y**
> - **Date Filter** = คอลัมน์กรองตามวันที่ (ตัว `Data Transaction Date`)
> - **Key Mapping** = คอลัมน์คีย์เชื่อม (join)
> - **# Select** = จำนวนคอลัมน์ที่เลือกแสดงผล

### C.1 Database — Product D
| Source | Date Filter | Key Mapping | # Select |
|--------|-------------|-------------|:--------:|
| `2_D001` | TRADE_DATE | TRADE_REF | 15 |
| `2_D002` | Data Set Date | FI Arrangement Number | 52 |
| `2_D003` | Date | Ref + CCY + Amount + CMF Code | 8 |
| `2_D004` | *(none)* | *(none)* | 5 |
| `2_D005` | *(none)* | *(none)* | 5 |

### C.2 Database — Product O
| Source | Date Filter | Key Mapping | # Select |
|--------|-------------|-------------|:--------:|
| `3_O001` | Arrangement Contract Date | FI Arrangement Number | 15 |
| `3_O002` | *(none)* | OptionNum | 13 |
| `3_O003` | Date | CompositeOptSeqNum | 6 |

### C.3 Database — Product C
| Source | Date Filter | Key Mapping | # Select |
|--------|-------------|-------------|:--------:|
| `4_C001` | *(none)* | Transaction No + DR/CR | 0 |
| `4_C002` | Settlement Date | Ccy + Notional | 2 |
| `4_C003` | Settlement Date | Currency + Accrued Interest | 2 |

### C.4 A1_System (main) — `1_A`
| Source | Date Filter | Key Mapping | # Select |
|--------|-------------|-------------|:--------:|
| `1_A001` | Data Set Date | FI Arrangement Number | 9 |
| `1_A002` | Data Set Date | Arrangement Number | 6 |
| `1_A003` | Data Set Date | FI Arrangement Number | 52 |
| `1_A004` | Data Set Date | Arr Number | 6 |
| `1_A005` | Data Set Date | FI Arrangement Number | 9 |
| `1_A006` | Data Set Date | FI Arrangement Number | 5 |
| `1_A007` | Data Set Date | FI Arrangement Number | 6 |
| `1_A008` | Data Set Date | Reference Transaction Number | 6 |
| `1_A009` | Data Set Date | FI Arrangement Number + CMF CODE | 12 |

### C.5 A1_System (ERROR output) — `1_B`
> ทุกไฟล์มีคอลัมน์ Select = `[Key, ERROR Message]` | Date Filter = *(none)*

| Source | Key Mapping |
|--------|-------------|
| `1_B001` | Arrangement Number |
| `1_B002` | Fi Arrangement Number |
| `1_B003` | Fi Arrangement Number |
| `1_B004` | Arrangement Number |
| `1_B005` | Fi Arrangement Number |
| `1_B006` | Fi Arrangement Number |
| `1_B007` | Fi Arrangement Number |
| `1_B008` | Ref.No. |
| `1_B009` | Fi Arrangement Number |

---

## หมวด D — Reconcile Rules (กฎการจับคู่)

> Default ที่ตั้งไว้ในเครื่องมือ — **ช่องที่มี ⚠️ = รอผู้ใช้ยืนยัน**

### D.1 Match Pattern — 3 แบบ
| แบบ | ชื่อ | คำอธิบาย |
|:--:|------|----------|
| **1** | Direct map | จับคู่ตรงด้วย Match Key แล้วเทียบ Compare Fields |
| **2** | Direct + Condition | กรองเงื่อนไขก่อน (Product/Transaction Type, Settle status) แล้วจึงจับคู่ |
| **3** | Map + Back-check | จับคู่ผ่านสายโซ่ แล้วย้อนกลับไปตรวจ Main Source (chain 3-4 ขั้น) |

### D.2 ตารางกฎต่อ Set (default)
| Set ID | Pattern | Match Key Fields | Compare Fields | Condition / Back-check |
|--------|:------:|------------------|----------------|------------------------|
| D1-1 | 2 ⚠️ | Ref (TRADE_REF=FI Arr No) | CCY + Amount + Date | PRODUCT_TYPE in scope |
| D1-2 | 1 | Ref (FI Arr No) | CCY + Buy/Sell Amount + Date | — |
| D1-3 | 2 ⚠️ | Ref (FI Arr No) | CCY + Original Amount + Date | Leg Type (1,2) |
| D1-4 | 1 | Ref (Arr Number) | CCY + Original Amount + Date | — |
| D2-5 | 3 ⚠️ | Ref (FI Arr No) | CCY + Amount + Date | back → `2_D001` |
| D3-6 | 2 ⚠️ | Ref + CCY + Amount | CCY + Amount + Date | Type / Settle status |
| D3-7 | 2 ⚠️ | Ref + CCY + Amount | CCY + Amount + Date | Type / Settle status |
| D3-8 | 2 ⚠️ | Ref + CCY + Amount | CCY + Amount + Date | Type / Settle status |
| D3-9 | 2 ⚠️ | Ref + CCY + Amount | CCY + Amount + Date | Type / Settle status |
| D3-10 | 2 ⚠️ | Ref + CCY + Amount | CCY + Amount + Date | Transaction Type |
| D3-11 | 2 ⚠️ | Ref + Type + CCY + Amount | CCY + Amount + Date | Type |
| D3-12 | 3 ⚠️ | Ref (CMF/Trans Ref) | CCY + Current Principal | back → `2_D004` + `2_D005` |
| O1-1 | 1 | Ref (FI Arr No) | CCY + Premium Amount + Date | — |
| O1-2 | 1 | Ref (FI Arr No) | CCY + Buy/Sell Amount + Date | — |
| O1-3 | 3 ⚠️ | Ref (FI Arr No / OptSeqNum) | CCY + Amount + Premium | back → `3_O003` |
| O1-4 | 1 | Ref (Arr Number) | CCY + Original Amount + Date | — |
| O2-5 | 2 ⚠️ | Ref (OptionNum) | CCY + Amount | Expired/Exercised status |
| O3-6 | 1 | Ref (OptSeqNum) | Premium CCY + Premium Amount + Date | — |
| C1-1 | 2 ⚠️ | Ref (Transaction No) + DR/CR | Amount | DR/CR split |
| C1-2 | 3 ⚠️ | Ref + Ccy + Notional | Ccy + Notional + Accrued Int | back → `4_C003` |
| C1-3 | 1 | Ref (Ref.No.) | CCY + Amount | — |

### D.3 Match Key forms (ฝั่ง n — ไม่รวมยอด)
- `Ref` อย่างเดียว
- `Ref + CCY + Amount`
- `Ref + Type + CCY + Amount`

---

## หมวด E — มาตรฐานคำศัพท์ & DSL

### E.1 Key Prep DSL (เตรียมคีย์ก่อน mapping)
| คำสั่ง | ความหมาย | ตัวอย่าง |
|--------|----------|----------|
| `AS_IS` | ใช้ค่าดิบ | `C0001` → `C0001` |
| `LEFT(n)` / `RIGHT(n)` | ตัด n ตัวซ้าย/ขวา | — |
| `MID(start,len)` | ตัดช่วงกลาง | — |
| `AFTER("x")` / `BEFORE("x")` | เอาส่วนหลัง/หน้าตัวอักษร x | — |
| `REGEX("pattern")` | จับด้วย regex | `KT20260530C0001` → `REGEX("C\d+")` → `C0001` |
| `CONCAT(f1,f2)` | รวมหลาย field | `Ref + CCY` |

### E.2 สถานะผล (Status)
| สถานะ | สี | ความหมาย |
|--------|----|----------|
| `OK / Matched` | 🟢 เขียว | ตรงกันทุก field |
| `MISSING in A1` | 🔴 แดง | มีใน DB ไม่มีใน A1 → บันทึกเพิ่มที่ A1 |
| `EXTRA in A1` | 🔴 แดง | มีใน A1 ไม่มีใน DB → ตรวจ/ลบที่ A1 |
| `VALUE MISMATCH` | 🟠 ส้ม | Key ตรง แต่ CCY/Amount/Date ต่าง → แก้ค่าที่ A1 |

### E.3 อภิธานศัพท์
| คำ | ความหมาย |
|----|----------|
| **A1_System** | ระบบหลักที่บันทึก Transaction (ไฟล์ `1_A`/`1_B`) — ตัวที่ถูกตรวจ |
| **Product Database** | ฐานข้อมูลต้นทางแยกตาม Product D/O/C (ไฟล์ `2_`/`3_`/`4_`) |
| **Set** | 1 ชุดการ Reconcile (1 สายโซ่ source) |
| **Trans.Map** | Cardinality เช่น `1:1:n` = 1 ต่อ 1 ต่อ หลาย |
| **Data Transaction Date** | วันที่ Transaction ที่ผู้ใช้ระบุเพื่อตรวจ (ตัวกรองหลัก) |
| **Match Key** | คีย์ที่ใช้จับคู่ระหว่าง DB ↔ A1 |
| **ERROR Message** | ข้อความผลการ Reconcile ที่นำกลับไปแก้ที่ A1 (เก็บในไฟล์ `1_B`) |

---

### ⧖ รายการรอยืนยัน (ดูชีท `Open_Questions` ในไฟล์ Excel)
1. **Pattern/Compare ต่อ Set** (โดยเฉพาะที่ติด ⚠️) — โปรดยืนยัน/แก้
2. **Key Prep rule จริงต่อ Source** — ต้องการรูปแบบคีย์ดิบจริงของไฟล์ที่ต้อง prep
3. **Pattern 3 Back-check** — ระบุว่าย้อนไปตรวจ field ใดที่ Main Source
4. **ข้อมูลตัวอย่างจริง** 1-2 ไฟล์ เพื่อ map คอลัมน์ + เขียน engine ให้รันครบทุก Set

---
*สร้างจาก `MockUp_File.xlsx` — ใช้คู่กับ `Part_Reconcile_Tool_Prototype.xlsx`*
