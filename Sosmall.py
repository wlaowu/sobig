import os
import requests
import zipfile
import subprocess
from art import text2art
from tabulate import tabulate
import requests
from tqdm import tqdm

# SDelete 下载链接和保存路径
sdelete_url = 'https://download.sysinternals.com/files/SDelete.zip'
user_profile = os.getenv('USERPROFILE')  # 获取用户目录路径
sdelete_zip_path = os.path.join(user_profile, 'SDelete.zip')  # 保存到用户目录
sdelete_dir = os.path.join(user_profile, 'SDelete')  # 解压到用户目录
sdelete_executable_path = os.path.join(sdelete_dir, 'sdelete.exe')


def get_available_drives():
    """获取系统中所有存在的驱动器盘符."""
    return [f"{drive}:\\" for drive in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ' if os.path.exists(f"{drive}:\\")]


def download_sdelete(url, save_path):
    """下载 SDelete 工具，并显示下载进度条."""

    print("⬇️正在下载 SDelete...")
    try:
        response = requests.get(url, stream=True)
        response.raise_for_status()  # 检查 HTTP 错误

        # 获取文件总大小
        total_size = int(response.headers.get('content-length', 0))
        chunk_size = 8192  # 每次读取的块大小

        # 使用 tqdm 显示进度条
        with open(save_path, 'wb') as file, tqdm(
                total=total_size, unit='B', unit_scale=True, desc=save_path, ncols=80) as pbar:
            for chunk in response.iter_content(chunk_size=chunk_size):
                file.write(chunk)
                pbar.update(len(chunk))  # 更新进度条

        print("SDelete下载完成.🔚")
    except requests.exceptions.RequestException as e:
        print(f"下载 SDelete 时出错: {e}")


def extract_sdelete(zip_path, extract_to):
    """解压 SDelete 工具，并显示解压进度条."""
    print()
    print("⋙  正在解压 SDelete... ⋘")
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            # 获取压缩文件中所有文件的信息
            file_list = zip_ref.infolist()
            total_size = sum(file.file_size for file in file_list)  # 计算总大小

            with tqdm(total=total_size, unit='B', unit_scale=True, desc="解压中", ncols=80) as pbar:
                for file in file_list:
                    extracted_path = zip_ref.extract(file, extract_to)
                    pbar.update(file.file_size)

        print("Sdelete解压完成.✔️")
    except zipfile.BadZipFile as e:
        print(f"解压 SDelete 时出错: {e}")
    except Exception as e:
        print(f"解压过程中发生错误: {e}")


def print_banner(message, width=40):
    """打印兼容中英文的消息框."""
    # 计算消息的实际长度，处理中英文混合的情况
    message_length = sum(2 if '\u4e00' <= char <= '\u9fff' else 1 for char in message)

    # 设置边框宽度，如果消息长度大于指定宽度，调整宽度
    frame_width = max(message_length + 4, width)

    # 计算内边距
    side_padding = (frame_width - message_length - 2) // 2
    side_space = ' ' * side_padding

    # 打印框架
    border = "■" * frame_width
    print(border)
    print(f"■{side_space}{message}{side_space}■")
    print(border)


def print_drive_list(drives):
    """使用简单的文本框架打印磁盘列表."""
    print("+---------------+")
    print("| Drive Letter  |")
    print("+---------------+")
    for drive in drives:
        print(f"| {drive:<13} |")
    print("+---------------+")


def run_sdelete(sdelete_path, disk):
    """运行 SDelete 清理指定磁盘的空闲空间."""

    # 自定义的提示信息
    message = f"Cleaning up free space on drive {disk}, please wait..."

    # 打印优美的提示信息
    print()
    print_banner(message)

    try:
        # 运行 sdelete 命令并打印输出
        subprocess.run(
            [sdelete_path, "-accepteula", "-p", "1", "-s", "-z", disk],
            shell=True
        )
    except Exception as e:
        print(f"运行 SDelete 时发生错误: {e}")


def main():
    # 创建目录并下载 SDelete 如果未存在
    try:
        if not os.path.exists(sdelete_executable_path):
            if not os.path.exists(sdelete_dir):
                os.makedirs(sdelete_dir, exist_ok=True)  # 创建目录
            download_sdelete(sdelete_url, sdelete_zip_path)
            extract_sdelete(sdelete_zip_path, sdelete_dir)
    except PermissionError as e:
        print(f"权限错误: {e}")
        print("你可能需要以管理员身份运行脚本或更改下载路径到你有写权限的目录.")
        return  # 退出程序
    except Exception as e:
        print(f"创建目录或下载 SDelete 时出错: {e}")
        return  # 退出程序

    # 获取所有盘符并打印
    drives = get_available_drives()
    print()
    print("🔃需要清理的磁盘如下：")
    print_drive_list(drives)

    # 对每个盘符运行 SDelete
    for drive in drives:
        run_sdelete(sdelete_executable_path, drive)

    print("\n磁盘清理完成.🔚")
    # 等待用户输入以保持窗口打开
    input("按 Enter 键退出...")


if __name__ == "__main__":
    main()
