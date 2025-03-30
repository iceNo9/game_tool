import os
import shutil
import sys
import argparse

def copy_directory(src_path):
    """ 拷贝 src_path 到当前目录 """
    home_dir = os.path.expanduser("~")  # 获取用户目录
    if not src_path.startswith(home_dir):
        raise ValueError("路径必须位于用户目录下")

    dst_path = os.path.join(os.getcwd(), os.path.basename(src_path))  # 目标路径
    shutil.copytree(src_path, dst_path, dirs_exist_ok=True)
    return dst_path

def generate_vbs(src_path):
    """ 生成 VBS 脚本，将拷贝的目录放回原位置 """
    home_dir = os.path.expanduser("~")
    rel_path = os.path.relpath(src_path, home_dir)  # 计算相对路径
    copied_dir = os.path.basename(src_path)
    copied_full_path = os.path.join(os.getcwd(), copied_dir).replace("/", "\\")  # 确保是完整路径

    vbs_content = f"""
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
homeDir = objShell.ExpandEnvironmentStrings("%USERPROFILE%")

fullPath = homeDir & "\\{rel_path.replace("/", "\\")}"

' 删除旧文件夹，防止冲突
If objFSO.FolderExists(fullPath) Then
    ' 弹出选择框，提示用户是否恢复存档
    response = MsgBox("即将重置存档,是否继续重置?如果不想覆盖存档请点否且手动备份"+fullPath, vbYesNo + vbQuestion, "恢复确认")
    If response = 6 Then ' 用户选择了“是”
        objFSO.DeleteFolder fullPath, True
    ElseIf response = 7 Then ' 用户选择了“否”
        WScript.Quit ' 退出脚本
    End If
End If

' 逐级检查并创建目录
CreateDirs fullPath

' 复制文件夹
If objFSO.FolderExists(fullPath) Then
    objFSO.CopyFolder "{copied_dir}", fullPath, True
    MsgBox("存档恢复成功")
Else
    MsgBox("存档恢复失败,请手动复制到用户目录下:{rel_path}")
End If

Sub CreateDirs(path)
    Dim folders, i, tempPath
    folders = Split(path, "\\")  ' 拆分路径
    tempPath = folders(0)  ' 例如 C:

    ' 逐级创建目录
    For i = 1 To UBound(folders)
        tempPath = tempPath & "\\" & folders(i)
        If Not objFSO.FolderExists(tempPath) Then
            objFSO.CreateFolder tempPath
        End If
    Next
End Sub
"""
    vbs_path = os.path.join(os.getcwd(), "restore.vbs")

    with open(vbs_path, "w", encoding="mbcs") as vbs_file:
        vbs_file.write(vbs_content.strip())

    return vbs_path

def main():
    parser = argparse.ArgumentParser(description="GameTool - 复制目录并生成 VBS 脚本")
    parser.add_argument("-g", "--generate", metavar="PATH", type=str, help="拷贝目录并生成 VBS 脚本", required=False)
    parser.add_argument("-V", "--version", action="version", version="GameTool 1.0.1")

    args = parser.parse_args()

    if args.generate:
        try:
            dst_path = copy_directory(args.generate)
            vbs_path = generate_vbs(args.generate)
            print(f"文件已复制到: {dst_path}")
            print(f"VBS 脚本已生成: {vbs_path}")
        except Exception as e:
            print(f"发生错误: {e}")
            sys.exit(1)

if __name__ == "__main__":
    main()
