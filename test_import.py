try:
    import vnstock
    print(f"Import vnstock thành công! Version: {vnstock.__version__}")
except Exception as e:
    print(f"Lỗi import: {e}")
