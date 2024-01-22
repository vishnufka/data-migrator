import config as cfg
import functions as fu
import os
import sys


def main():
    print(cfg.get_start_time())

    # if we need to hash anything we need them to be consistent
    if int(os.environ['PYTHONHASHSEED']) != 0:
        sys.exit()

    for filename in cfg.files_list:
        df = fu.read_excel(cfg.data_path + filename, cfg.header_row_index)
        df = fu.prepare_df(df, filename)
        df = fu.send_df_to_client_method(df, filename)
        fu.export_to_excel(df, filename)
        print(filename)
        print(cfg.get_end_time())
        print(cfg.get_run_time())

    fu.run_extra_client_methods()


if _name_ == "_main_":
    main()