import os
import tempfile
import CountCode
import sys

project_list = ['git clone ssh://wangxianchao@igerrit.storm:29418/Src/Group/Data/Code/Search',
                'git clone ssh://wangxianchao@igerrit.storm:29418/Src/Group/Data/Code/Road']

project_list = ['git clone ssh://wangxianchao@igerrit.storm:29418/Src/Group/Data/Code/Search']


def clone_project(project_list, path):
    pass


def traverse_branch(project_path):
    os.chdir(project_path)
    branch = tempfile.mktemp(suffix='_branches.txt')
    res = os.system("git branch -r > %s" % branch)
    branch_name_list = []
    if res == 0:
        with open(branch, 'r') as f:
            branch_name = f.readline()
            while branch_name:
                branch_name = branch_name.strip()
                branch_name = branch_name.strip('\n')
                if 'HEAD' in branch_name:
                    branch_name = f.readline()
                    continue
                branch_name_list.append(branch_name)
                branch_name = f.readline()
    else:
        return 1

    return branch_name_list


if __name__ == '__main__':
    project_dir = sys.argv[1]
    outputdir = sys.argv[2]

    out_put_path = os.path.abspath(outputdir)
    if not os.path.exists(out_put_path):
        os.mkdir(out_put_path)
    basefilepath = out_put_path

    # project_dir = "C:\\AProject\\16Tmap"
    origin_branches = traverse_branch(project_dir)
    for branch in origin_branches:
        print(branch, "\t", basefilepath + "_" + branch.replace('/', '_') + ".xlsx")
        CountCode.main(branch, "C:\\AProject\\16Tmap", basefilepath + "/" + branch.replace('/', '_') + ".xlsx")


