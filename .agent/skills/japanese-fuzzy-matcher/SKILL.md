---
name: japanese-fuzzy-matcher
description: Match Japanese e-commerce product titles against official catalog models using full-width/half-width normalization, knife type ontology, blade length extraction, and fuzzy similarity scoring.
---

# Japanese Fuzzy Matcher Skill

This skill provides algorithms and text processing pipelines to match Japanese product listings to official catalog codes when the explicit alphanumeric model code (e.g., `G-46`, `GST-B46`) is missing from the title.

## 🎯 When to Use This Skill

- To resolve products classified as 🟡 `CODE_NOT_FOUND (YELLOW)` in price comparison reports.
- Normalizing mixed Japanese typography (Zen-kaku 全角 vs Han-kaku 半角, Katakana vs Hiragana, Kanji).
- Extracting key product attributes (Blade Type, Length in cm/mm, Series: Standard vs IST vs PRO).

---

## 🔍 Text Normalization & Attribute Extraction

### 1. Unicode NFKC Normalization

Always normalize text with `unicodedata.normalize('NFKC', text)` to convert full-width numbers/letters (`１８ｃｍ`, `Ｇ－４６`) into standard ASCII (`18cm`, `G-46`).

```python
import re
import unicodedata
from typing import Dict, Optional, Tuple


def normalize_japanese_title(text: str) -> str:
    """Normalizes Unicode characters, full-width numbers, and whitespace."""
    if not text:
        return ""
    normalized = unicodedata.normalize("NFKC", text)
    # Collapse multiple spaces
    return re.sub(r"\s+", " ", normalized).strip()
```

---

### 2. Japanese Knife Domain Ontology & Attribute Extractor

```python
# Knife Types Mapping
KNIFE_TYPES: Dict[str, str] = {
    "三徳": "SANTOKU",
    "万能": "SANTOKU",
    "牛刀": "GYUTO",
    "シェフナイフ": "GYUTO",
    "ペティ": "PETTY",
    "ペティー": "PETTY",
    "小型": "PETTY",
    "パン切り": "BREAD",
    "ブレッド": "BREAD",
    "菜切り": "NAKIRI",
    "出刃": "DEBA",
    "刺身": "SASHIMI",
    "柳刃": "YANAGIBA",
    "筋引": "SUJIHIKI",
    "皮むき": "PEELING",
}


def extract_knife_attributes(
    title: str,
) -> Tuple[Optional[str], Optional[int]]:
    """Extracts knife type category and blade length in mm from title."""
    norm_title = normalize_japanese_title(title)

    # 1. Detect Knife Type
    detected_type: Optional[str] = None
    for keyword, category in KNIFE_TYPES.items():
        if keyword in norm_title:
            detected_type = category
            break

    # 2. Extract Blade Length (e.g., '18cm', '180mm', '18 センチ')
    detected_length_mm: Optional[int] = None
    cm_match = re.search(r"(\d+(?:\.\d+)?)\s*(?:cm|ｃｍ|センチ)", norm_title, re.I)
    if cm_match:
        detected_length_mm = int(float(cm_match.group(1)) * 10)
    else:
        mm_match = re.search(r"(\d+)\s*(?:mm|ｍｍ)", norm_title, re.I)
        if mm_match:
            detected_length_mm = int(mm_match.group(1))

    return detected_type, detected_length_mm
```

---

### 3. Knife Sets & Bundles Matching (`点セット` / `点set`)

When titles describe a set (e.g., `牛刀（16cm）2点セット`, `三徳 3点セット`):

- Extract set count: `2点セット` ➔ `set_count = 2`, `3点セット` ➔ `set_count = 3`, `4点セット` ➔ `set_count = 4`.
- Combine with knife category and blade length:
  - `牛刀(16cm) 2点セット` ➔ **`GST-A58`**
  - `牛刀(20cm) 2点セット` ➔ **`GST-A2`**
  - `牛刀(18cm) 3点セット` ➔ **`GST-B4`**
  - `牛刀(20cm) 3点セット` ➔ **`GST-B2`**
  - `牛刀(20cm) 4点セット` ➔ **`GST-C2`**
  - `三徳(18cm) 2点セット` ➔ **`GST-A46`**
  - `三徳(16cm) 2点セット` ➔ **`GST-A57`**
  - `三徳(18cm) 3点セット` ➔ **`GST-B46`**
  - `三徳(16cm) 3点セット` ➔ **`GST-B57`**
  - `三徳(18cm) 4点セット` ➔ **`GST-C46`**
  - `ペティーナイフ 2点セット` ➔ **`GST-AS3`**

---

### 4. Catalog Model Disambiguation Scoring

When matching a title against the catalog without a direct code:

1. **Series Match:** Check if title contains `IST` (Global-IST), `PRO`, or standard `GLOBAL`.
2. **Set vs Single Knife:** Determine if the item is a bundle (`set_count >= 2`) or individual knife.
3. **Type + Length Match:**
   - Single: `三徳 18cm` in standard series ➔ `G-46`
   - Set: `牛刀 16cm 2点セット` ➔ `GST-A58`
4. **Similarity Threshold:** Assign code with high confidence based on the multi-attribute tuple `(Series, Category, Length, SetCount)`.
