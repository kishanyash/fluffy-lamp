import requests

base_url = "https://bmpvcjbfeyvkkbvclwkb.supabase.co/storage/v1/object/public/charts/ETERNAL/"
timestamp = "20260217_145435"

candidates = [
    # Variations of name
    "price_chart", "chart_price", "price", "stock_price", "chart_stock", 
    "price_history", "chart_history", "price_trend", "chart_trend",
    "technical_chart", "chart_technical"
]

extensions = [".png", ".jpg", ".jpeg"]

with open("verification_result.txt", "w") as f:
    f.write("Checking URLs (Extended)...\n")
    found = False
    for c in candidates:
        for ext in extensions:
            filename = f"{c}_{timestamp}{ext}"
            url = f"{base_url}{filename}"
            try:
                r = requests.head(url)
                if r.status_code == 200:
                    f.write(f"FOUND: {url}\n")
                    found = True
                    break
            except:
                pass
        if found: break
    
    if not found:
         f.write("NO MATCH FOUND (Extended).\n")
