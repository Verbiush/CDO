
try:
    import selenium
    print("Selenium: OK")
except ImportError as e:
    print(f"Selenium: FAIL ({e})")

try:
    import webdriver_manager
    print("WebDriver Manager: OK")
except ImportError as e:
    print(f"WebDriver Manager: FAIL ({e})")
