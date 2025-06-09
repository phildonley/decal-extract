Small victory today... finally resolved the api key retrieval issue, created a fallback that asks the user to re-enter the key, just in case the file becomes corrupted. 

However, now we are back to the real problems which were cropping. So we need to relook at the failing crop methods and re-implement some things that we seem to have lose while trying to fix the switch from 
to the DocLibrary API

Load times are crazy fast. But now we need to fix the crop logic. 

C:\Users\Phillip.Donley\DECAL EXTRACT>python decalextract.py
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
libpng warning: iCCP: known incorrect sRGB profile
· Loaded 12 template sets for corner detection
[0] ➡️ Processing part=65417GT, TMS=19483786
[DEBUG] DNS lookup succeeded for hal4ecrr1k.execute-api.us-east-1.amazonaws.com → 3.218.4.17
[OK] Downloaded PDF → C:/Users/Phillip.Donley/DECAL EXTRACT\decal_output_06092025\temp_pdfs\65417GT_1749499496.pdf
    · PDF downloaded → C:/Users/Phillip.Donley/DECAL EXTRACT\decal_output_06092025\temp_pdfs\65417GT_1749499496.pdf
   · Attempting crop-mark → bracket → union-of-all-ink…
Traceback (most recent call last):
  File "C:\Users\Phillip.Donley\DECAL EXTRACT\decalextract.py", line 1185, in <module>
    main(sheet, out_root, seq=105)
  File "C:\Users\Phillip.Donley\DECAL EXTRACT\decalextract.py", line 1034, in main
    cv2.imwrite(out_jpg, crop_img)
                         ^^^^^^^^
NameError: name 'crop_img' is not defined

C:\Users\Phillip.Donley\DECAL EXTRACT>
