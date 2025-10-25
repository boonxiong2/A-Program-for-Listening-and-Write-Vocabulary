import win32com.client
from time import sleep
import os

# 初始化语音引擎
speaker = win32com.client.Dispatch("SAPI.SpVoice")
voices = speaker.GetVoices()

# 程序版本信息
VERSION = "1.0.0.20251025.1-Release"
RELEASE_DATE = "2025-10-25"
AUTHOR = "英语听写助手开发者:熊羿然"
CONTACT = "16048561@qq.com"


def show_version_info():
    """显示版本信息"""
    print(f"\n{'=' * 40}")
    print(f"英语听写助手终端版 v{VERSION}")
    print(f"发布日期: {RELEASE_DATE}")
    print(f"开发者: {AUTHOR}")
    print(f"联系方式(E-mail): {CONTACT}")
    print(f"{'=' * 40}")


def main():
    """主程序"""
    # 显示版本信息
    show_version_info()

    # 获取用户输入
    usr_input = int(input("（默认六年级上册沪教版）听写第几单元？(1-10)（11是自定义）:"))

    # 处理自定义单元
    if usr_input == 11:
        filename = input("请输入自定义单词表文件名（文件路径，如果在同一文件夹下，直接写文件名））: ")
        if not os.path.exists(filename):
            print(f"没有找到文件: {filename}")
            speaker.Speak(f"没有找到文件 {filename}")
            return
    else:
        filename = f'u{usr_input}.txt'

    # 确保单元文件存在
    if not os.path.exists(filename):
        print(f"没有找到单元 {usr_input} 的文件")
        speaker.Speak(f"没有找到单元 {usr_input} 的文件")
        return

    # 检查文件是否为空
    if os.path.getsize(filename) == 0:
        print(f"文件内容为空: {filename}")
        speaker.Speak(f"文件内容为空")
        return

    print(f"正在听写: {filename}")
    speaker.Speak("正在听写")
    speaker.Voice = voices.Item(1)  # 英文语音
    speaker.Speak(f"Unit {usr_input if usr_input != 7 else 'Custom'}")

    # 读取单元单词
    with open(filename, 'r', encoding='utf-8') as r:
        words = [line.strip() for line in r if line.strip()]

    # 开始听写

    for i, word in enumerate(words, 1):
        print(f"\n第 {i} 个单词", end="\n\n")

        # 用英文语音朗读提示和单词
        try:
            speaker.Voice = voices.Item(1)  # 英文语音
            speaker.Speak(f"Number {i}")
            sleep(0.5)
            speaker.Speak(word)
            sleep(2.5)
            speaker.Speak(word)
            sleep(1)
        except KeyboardInterrupt:
            input("程序中止，按下回车键退出...\n")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"程序出错: {e}")
        speaker.Speak("程序出错了，请检查设置")
    finally:
        input("\n按回车键退出程序...")