import os

image_file_dir = os.path.join(os.getcwd(), "public/stocklist")
image_dir_list = os.listdir(image_file_dir)


image_dir_list.sort()
print(image_dir_list)

for i, file in enumerate(image_dir_list):
    os.rename(
        f"{image_file_dir}/{file}",
        f"{image_file_dir}/stocklist_forklift_sample_{i+2}.jpg",
    )
