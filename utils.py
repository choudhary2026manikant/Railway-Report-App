def process_image(img_bytes):
    try:
        img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
        img = ImageOps.fit(img, (800, 600), Image.Resampling.LANCZOS)
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(1.05)
        
        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format='JPEG', quality=85, optimize=True)
        img_byte_arr.seek(0)
        return img_byte_arr
    except Exception as e:
        # Fallback agar image corrupt ho
        blank_img = Image.new('RGB', (800, 600), color=(200, 200, 200))
        img_byte_arr = io.BytesIO()
        blank_img.save(img_byte_arr, format='JPEG')
        img_byte_arr.seek(0)
        return img_byte_arr
