import os
from glob import glob

current_dir = os.getcwd()
re_current_dir = os.path.join(current_dir,"public/stocklist")
image_dir = os.listdir(re_current_dir)


image_dir.sort()
print(image_dir)

for i, file in enumerate(image_dir):
    os.rename(f"{re_current_dir}/{file}",f"{re_current_dir}/stocklist_forklift_sample_{i+2}.jpg")
