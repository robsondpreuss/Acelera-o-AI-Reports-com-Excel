# Porsche Sales Sanitization Schema

> Columns Data map

## Canonical input columns

| Column | Required | Description | Example raw values |
|---|---:|---|---|
| `sale_id` | yes | Unique sale record ID | `1`, `86` |
| `sale_date` | yes | Raw sales date | `2024-13-05`, `02/05/24`, `July 15th, 2024` |
| `customer_name` | yes | Customer name | `John  Smith`, `MARY johnson` |
| `porsche_model` | yes | Porsche model and trim | `911 Carrera`, `Macan Electric` |
| `model_year` | yes | Vehicle model year | `2023`, `20-20`, `two thousand twenty two` |
| `sale_price` | yes | Sale price in messy USD formats | `$985,000.00`, `188k USD` |
| `vehicle_mileage` | yes | Vehicle mileage in messy miles/km formats | `12,500 miles`, `KM 15.200`, `new` |
| `payment_method` | yes | Payment method | `CreditCard`, `bank-transfer`, `crypto payment` |
| `city` | yes | City | `new york`, `San Jose` |
| `state` | yes | US state name or abbreviation | `ny`, `California` |
| `salesperson` | yes | Salesperson name | `William  Clark`, `SARAH lee` |
| `delivery_status` | yes | Delivery status | `Delivered!!!`, `in-transit`, `DELIVERD` |

## Sanitized output columns

| Source | Output | Format |
|---|---|---|
| `sale_date` | `SaleDateSanitized` | `YYYY-MM-DD` or `INVALID` |
| `porsche_model` | `PorscheModelSanitized` | Canonical Porsche model/trim label |
| `model_year` | `ModelYearSanitized` | Four-digit year or `INVALID` |
| `sale_price` | `SalesPriceSanitized` | USD amount as decimal string with two decimals |
| `vehicle_mileage` | `VehicleMileageSanitized` | Integer miles |
| `payment_method` | `PayMethodSanitized` | Controlled label |
| `city` | `CitySanitized` | Title-cased city |
| `state` | `StateSanitized` | USPS two-letter code or `INVALID` |
| `delivery_status` | `DeliveryStatusSanitized` | Controlled label |

## Sanitization rules

### Dates

Normalize to ISO `YYYY-MM-DD`.

Accept common formats:

- `YYYY-MM-DD`
- `YYYY/MM/DD`
- `YYYY.MM.DD`
- `MM/DD/YYYY`
- `MM/DD/YY`
- `MM-DD-YY`
- `Month DDth, YYYY`
- `Mon DDth YYYY`

Invalid calendar dates such as `2024-13-05`, `2024-02-30`, `April 31st, 2024`, and `2027-06-40` must become `INVALID`.

### Porsche models

Normalize model names to canonical title case. Preserve valid trims such as:

- `911 Carrera`
- `911 Carrera S`
- `911 Carrera GTS`
- `911 Turbo`
- `911 Turbo S`
- `911 GT3`
- `911 GT3 RS`
- `911 Dakar`
- `911 Targa 4`
- `911 Targa 4S`
- `718 Cayman`
- `718 Cayman S`
- `718 Cayman GT4 RS`
- `718 Boxster`
- `718 Boxster GTS`
- `718 Spyder RS`
- `Cayenne`
- `Cayenne S`
- `Cayenne Coupe`
- `Cayenne E-Hybrid`
- `Cayenne Turbo`
- `Cayenne Turbo GT`
- `Macan`
- `Macan S`
- `Macan T`
- `Macan GTS`
- `Macan Electric`
- `Panamera`
- `Panamera 4`
- `Panamera 4S`
- `Panamera Turbo`
- `Panamera Turbo S`
- `Panamera 4 E-Hybrid`
- `Taycan`
- `Taycan 4S`
- `Taycan GTS`
- `Taycan Turbo`
- `Taycan Turbo S`
- `Taycan Cross Turismo`

Unknown models should be title-cased rather than dropped.

### Model year

Normalize to a four-digit year. Convert common text forms like `twenty twenty four` and formats like `20-24` or `20 24`. Values outside `1990` to `2035` should become `INVALID`.

### Sales price

Normalize to a decimal USD amount without symbols, commas, or text.

Examples:

- `$985,000.00` -> `985000.00`
- `870000 USD` -> `870000.00`
- `$645k` -> `645000.00`
- `188k USD` -> `188000.00`
- `eighty two thousand USD` -> `82000.00`
- `two hundred thousand USD` -> `200000.00`

### Vehicle mileage

Normalize to integer miles.

Examples:

- `12,500 miles` -> `12500`
- `Miles: 15,200` -> `15200`
- `zero miles`, `new`, `0 mi`, `new car` -> `0`
- `twelve thousand miles` -> `12000`

If a value is explicitly labeled `KM`, convert kilometers to miles using `1 km = 0.621371 miles`, rounded to the nearest integer.

### Payment method

Normalize to one of:

- `Credit Card`
- `Debit Card`
- `Bank Transfer`
- `Wire Transfer`
- `Financing`
- `Lease`
- `Cash`
- `ACH Payment`
- `Crypto Payment`

Unknown payment methods should be title-cased.

### City

Normalize city names to title case and preserve known punctuation such as `St. Louis`.

### State

Normalize US states to USPS two-letter codes. Convert both names and abbreviations, case-insensitively.

Examples:

- `California` -> `CA`
- `california` -> `CA`
- `ca` -> `CA`
- `New York` -> `NY`

Unknown states should become `INVALID`.

### Delivery status

Normalize to one of:

- `Delivered`
- `Pending`
- `In Transit`
- `Cancelled`
- `Awaiting Delivery`
- `Awaiting Pickup`
- `Pending Approval`
- `Pending Review`
- `Shipped`
- `Awaiting Review`

Handle punctuation, case, hyphenation, and common typo `DELIVERD`.

## Quality checks

After normalization, verify:

- No original columns were removed.
- Sanitized columns appear immediately after their source columns.
- Dates are ISO format or `INVALID`.
- State values are two-letter USPS codes or `INVALID`.
- Mileage is integer miles.
- Price is numeric with two decimals.
- No new blank values are introduced in sanitized columns; use `INVALID` or normalized defaults instead.

## Invalid handling

Use `INVALID` when a value cannot be normalized safely. Never leave a sanitized field blank.