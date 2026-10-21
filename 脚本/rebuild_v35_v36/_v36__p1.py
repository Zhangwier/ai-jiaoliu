import io, re
for f, old, new in [
  ("app2.py", "from app import sec, term_rows, num_list, b_top, b_matrix, a1, a2, a_index, BRAND, CH",
              "from app import sec, term_rows, num_list, b_top, b_matrix, a1, a2, a_index, BRAND, CH, hd"),
  ("app3.py", "from app import sec, term_rows, num_list, BRAND, CH",
              "from app import sec, term_rows, num_list, BRAND, CH, hd"),
  ("app4.py", "from app import sec, term_rows, num_list, CH",
              "from app import sec, term_rows, num_list, CH, hd"),
  ("app5.py", "from app import sec, CH", "from app import sec, CH, hd"),
]:
    s = io.open(f, encoding="utf-8").read()
    if old not in s:
        print("MISS", f); continue
    io.open(f, "w", encoding="utf-8").write(s.replace(old, new, 1))
    print("ok", f)
