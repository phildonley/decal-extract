C:\Users\Phillip.Donley\DECAL EXTRACT>python decalextract.py
Traceback (most recent call last):
  File "C:\Users\Phillip.Donley\DECAL EXTRACT\decalextract.py", line 1220, in <module>
    main(sheet, out_root, seq=105)
  File "C:\Users\Phillip.Donley\DECAL EXTRACT\decalextract.py", line 982, in main
    api_key = get_valid_api_key()    # this also sets DecalExtract_helper.API_KEY internally
              ^^^^^^^^^^^^^^^^^^^
  File "C:\Users\Phillip.Donley\DECAL EXTRACT\decalextract.py", line 50, in get_valid_api_key
    if os.path.exists(KEY_FILE):
                      ^^^^^^^^
NameError: name 'KEY_FILE' is not defined

C:\Users\Phillip.Donley\DECAL EXTRACT>
