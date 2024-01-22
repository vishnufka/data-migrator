
import datetime as dt
import clientname.config.client_config as client_cfg

client_folder_name = "clientname"

start_time = None
end_time = None
today_timestamp = dt.datetime.now().strftime("%Y%m%d")


def get_start_time():
    global start_time
    start_time = dt.datetime.now()
    return start_time.strftime("%Y-%m-%d %H:%M:%S")


def get_end_time():
    global end_time
    end_time = dt.datetime.now()
    return end_time.strftime("%Y-%m-%d %H:%M:%S")


def get_run_time():
    return end_time - start_time


# client-specific data

# used in fct.py
data_path = client_cfg.data_path
header_row_index = client_cfg.header_row_index
files_list = client_cfg.files_list

# used in functions.py
base_path = client_cfg.base_path
output_path = client_cfg.output_path
data_mapping_dict = client_cfg.data_mapping_dict
mapping_dict = client_cfg.mapping_dict
