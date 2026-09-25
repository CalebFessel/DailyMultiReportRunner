import traumasoft_api, os
for k in ("PAYCOR_ACCESS_TOKEN", "PAYCOR_REFRESH_TOKEN"):
    v = os.environ.get(k, "")
    print(f"{k:22} {'SET len=' + str(len(v)) if v.strip() else '(not set)'}")
