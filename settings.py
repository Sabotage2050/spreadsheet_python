import os
from os.path import join, dirname
from dotenv import load_dotenv

load_dotenv(verbose=True)

dotenv_path = join(dirname(__file__), ".env")

load_dotenv(dotenv_path)

TEST = os.getenv('TEST').split(',')
# print(f"env_list:{env_list}")
# print('ListItem is ', len(env_list))
# 
# for env in env_list:
    # print(env)
