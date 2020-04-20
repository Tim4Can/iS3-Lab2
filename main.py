import os
import sys
import config
import importlib


def check_datatype(data_type):
    for dt in config.datatypes:
        if dt in data_type.lower():
            return config.datatypes[dt]
    return None


def traverse(input_path, output_path, project, datatype=None):
    if (project, datatype) in config.tasks:
        processor_class = importlib.import_module(config.tasks[(project, datatype)])
        processor = processor_class.Processor()
        if hasattr(processor, "run"):
            run = getattr(processor, 'run')
            run(input_path, output_path)

    else:
        if datatype is None:
            datatype = ""
        print("当前不支持以下数据格式的操作：" + project + " " + datatype)
        return


def main(input_path, output_path):
    for dir_name in os.listdir(input_path):
        if os.path.isdir(dir_name):
            continue
        if dir_name not in config.projects.keys():
            print("文件夹：" + dir_name + " 不在目标列表中。")
            continue
        need_dt = config.projects[dir_name]
        if need_dt:
            for data_type in os.listdir(os.path.join(input_dir, dir_name)):
                dt = check_datatype(data_type)
                traverse(os.path.join(input_dir, dir_name, data_type), output_path, dir_name, dt)
        else:
            traverse(os.path.join(input_dir, dir_name), output_path, dir_name)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise ValueError("输入参数个数不正确！")
    input_dir = sys.argv[1]
    output_dir = sys.argv[2]
    if not os.path.exists(input_dir):
        raise FileNotFoundError("输入文件夹 " + input_dir + " 不存在！")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    main(input_dir, output_dir)
