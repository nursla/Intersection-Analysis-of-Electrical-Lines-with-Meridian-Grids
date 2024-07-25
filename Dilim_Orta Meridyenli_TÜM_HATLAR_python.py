#!/usr/bin/env python
# coding: utf-8

# In[31]:


#NURSILAÖZKAN
#25/07/2024


import geopandas as gpd
import pandas as pd
import os

electric_lines_path = r'C:\Users\zag7ols\Desktop\teiaş hatlar'
paftalar_path = r'C:\Users\zag7ols\Desktop\TM'
output_path = r'C:\Users\zag7ols\Desktop\electric_lines_projected.xlsx'

electric_lines_files = [os.path.join(electric_lines_path, f) for f in os.listdir(electric_lines_path) if f.endswith('.shp')]

paftalar_files = [os.path.join(paftalar_path, f) for f in os.listdir(paftalar_path) if f.endswith('.shp')]

electric_lines = gpd.read_file(electric_lines_files[0])

meridians = [27, 30, 33, 36, 39, 42, 45]

column_order = [
    'ID', 'hatadi', 'AsamaID', 'hat_id', 'DOM', 'bolge', 'hatno', 'hattip_id', 
    'hatozel_id', 'ikzkure_id', 'teskur_id', 'gunceltar', 'rapkur_id', 'altkur_id', 
    'aciklama', 'gerilim', 'kayityap', 'editleyen', 'kayittar', 'edittar', 'uzunluk', 
    'domproj', 'hattip', 'BolgeMudur', 'hatozel', 'ikzkure', 'teskur', 'rapkur', 
    'altkur', 'ProjeID', 'elektrikHa', 'elektrikLi', 'sebekeTipi', 'isletmeGer', 
    'azamiGeril', 'NAME', 'LAYER', 'MAP_NAME', 'PERIMETER', 'ENCLOSED_A', 'LENGTH', 
    'WIDTH', 'geometry'
]

with pd.ExcelWriter(output_path) as writer:
    for meridian in meridians:
        
        pafta_file = [f for f in paftalar_files if f'_{meridian}' in f]
        if not pafta_file:
            print(f"Meridian {meridian} için pafta bulunamadı.")
            continue

        paftalar = gpd.read_file(pafta_file[0])

        if electric_lines.crs != paftalar.crs:
            electric_lines = electric_lines.to_crs(paftalar.crs)

        intersected = gpd.overlay(electric_lines, paftalar, how='intersection')

        intersected = intersected.reindex(columns=column_order)

        # Sonuçları Excel'e yaz
        intersected.to_excel(writer, sheet_name=f'Meridian_{meridian}', index=False)

        print(f"Meridian {meridian} için kesişim sonuçları kaydedildi.")

print("Tüm kesişimler tek bir Excel dosyasına toplandı.")


#NURSILAÖZKAN
#25/07/2024


# In[ ]:




