C:\Users\Phillip.Donley\DECAL EXTRACT>python decalextract.py
Traceback (most recent call last):
  File "C:\Users\Phillip.Donley\DECAL EXTRACT\decalextract.py", line 1036, in <module>
    x0, y0, x1, y1 = select_best_crop_box(img, template_sets, expected_ratio)
                                          ^^^
NameError: name 'img' is not defined

During handling of the above exception, another exception occurred:

Traceback (most recent call last):
  File "C:\Users\Phillip.Donley\DECAL EXTRACT\decalextract.py", line 1041, in <module>
    H, W = img.shape[:2]
           ^^^
NameError: name 'img' is not defined

C:\Users\Phillip.Donley\DECAL EXTRACT>
