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
[DEBUG] DNS lookup succeeded for hal4ecrr1k.execute-api.us-east-1.amazonaws.com → 54.225.234.245
[OK] Downloaded PDF → C:/Users/Phillip.Donley/DECAL EXTRACT\decal_output_06092025\temp_pdfs\65417GT_1749508107.pdf
    · PDF downloaded → C:/Users/Phillip.Donley/DECAL EXTRACT\decal_output_06092025\temp_pdfs\65417GT_1749508107.pdf
   · Attempting bracket crop…
   · Bracket crop failed; trying template-corner crop…
   · Template-corner failed; trying nearby-blob grouping…
   · Nearby-blob failed; trying horizontal-aligned union…
   · [OK] Using horizontal-union crop: (1403, 896, 1897, 2010)
   · Final crop size: 494×1114
   · Writing JPEG → C:/Users/Phillip.Donley/DECAL EXTRACT\decal_output_06092025\images\19483786.65417GT.105.jpg
Traceback (most recent call last):
  File "C:\Users\Phillip.Donley\DECAL EXTRACT\decalextract.py", line 1234, in <module>
    main(sheet, out_root, seq=105)
  File "C:\Users\Phillip.Donley\DECAL EXTRACT\decalextract.py", line 1091, in main
    vol  = h_in * w_in * THICKNESS_IN
           ^^^^
NameError: name 'h_in' is not defined. Did you mean: 'h_in2'?

C:\Users\Phillip.Donley\DECAL EXTRACT>
