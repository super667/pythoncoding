import sys, re
import os
import tempfile
import pandas
import datetime
start_date = datetime.datetime.strptime('2019-07-01', "%Y-%m-%d")
end_date = datetime.datetime.now()


class GitInfo:
    def __init__(self):
        self.commitid = ""
        self.name = ""
        self.author = ""
        self.datetime = ""
        self.time = ""
        self.change_count = 0
        self.add = 0
        self.delete = 0
        self.change = 0
        self.changeID = ""


def main(origin_branch_name, project_dir, outputfile):
    pwd = os.getcwd()
    if not os.path.isdir(project_dir):
        print("%s doest not exist or is not a directory." % project_dir)
    os.chdir(project_dir)
    # ## get branch
    branch = tempfile.mktemp(suffix='_branches.txt')
    res = os.system("git branch > %s" % branch)
    # branch_name = 'master'
    # if res == 0:
    #     with open(branch, 'r') as f:
    #         branch_name = f.readline()
    #         if branch_name:
    #             branch_name = branch_name[2:]
    #             branch_name = branch_name.strip('\n')
    # else:
    #     return 1
    remote = tempfile.mktemp(suffix='_remote.txt')
    res = os.system("git remote -v > %s" % remote)
    path = ''
    if res == 0:
        with open(remote, 'r') as f:
            path = f.readline()
            if path:
                path = path.split('@')[1]
                path = path.strip('\n')
                # (fetch)
                path = path[:-7]
                path = path.strip(' ')
                if path.endswith('.git'):
                    path = path[:-4]
    # ## config git
    print("git config log.date iso")
    res = os.system("git config log.date iso")
    if res != 0:
        return 2
    temp1 = tempfile.mktemp(suffix='_1.txt')
    print("git log %s --shortstat > %s" % (origin_branch_name, temp1))
    res = os.system("git log %s --shortstat > %s" % (origin_branch_name, temp1))
    if res != 0:
        return 3
    temp2 = tempfile.mktemp(suffix='_2.txt')
    print("output %s" % temp2)
    count_main(temp1, temp2)
    df = pandas.read_csv(temp2, sep='\t')
    # df.insert(0, 'changeid', '')
    df.insert(2, 'branch', origin_branch_name)
    df.insert(3, 'path', path)
    os.chdir(pwd)
    print("output %s" % outputfile)
    df.to_excel(outputfile, sheet_name='code提交统计', index=False)
    print("Clear Temp files...")
    os.remove(temp1)
    os.remove(temp2)
    os.remove(branch)
    os.remove(remote)


def count_main(test_file_path, outputfile):
    # start_date = datetime.datetime.strptime('2019-07-01', format="%Y-%m-%d")
    with open(outputfile, mode='w') as op:
        header = ["changeid", "commitid",
                  "name", "author",
                  "datetime", "change_count",
                  "add", "del"]
        op.write('\t'.join(header) + '\n')
        test_cases = open(test_file_path, 'r', encoding='utf-8').read().split("\n\ncommit ")
        for testcase in test_cases:
            if (testcase.find(" files changed") != -1 or testcase.find(
                    " file changed") != -1 and testcase.strip() != ""):
                gitinfo = GitInfo()
                testcaselist = testcase.split("\n")
                gitinfo.commitid = testcaselist[0]
                gitinfo.name = re.split(' +', testcaselist[1])[1]
                gitinfo.author = re.split(' +', testcaselist[1])[2][1:-1]
                gitinfo.datetime = re.split(' +', testcaselist[2])[1]
                gitinfo.time = re.split(' +', testcaselist[2])[2]
                for caselist in testcaselist:
                    if (len(re.split(' +', caselist)) == 3 and
                            re.split(' +', caselist)[1] == "Change-Id:"):
                        gitinfo.changeID = re.split(' +', caselist)[-1]
                # if gitinfo.changeID == "":
                #     gitinfo.changeID = "no"
                if testcaselist[-1] != "":
                    changelist = re.split(',', testcaselist[-1])
                    gitinfo.change_count = int(changelist[0].strip().split(" ")[0])
                    if len(changelist) == 3:
                        gitinfo.add = int(changelist[1].strip().split(" ")[0])
                        gitinfo.delete = int(changelist[2].strip().split(" ")[0])
                    elif changelist[1].strip().split(" ")[1] == "insertions(+)":
                        gitinfo.add = int(changelist[1].strip().split(" ")[0])
                    else:
                        gitinfo.delete = int(changelist[1].strip().split(" ")[0])
                    # gitinfo.change = gitinfo.add + gitinfo.delete
                    # year = int(gitinfo.datetime.strip().split("-")[0])
                    # month = int(gitinfo.datetime.strip().split("-")[1])
                    commit_date = datetime.datetime.strptime(gitinfo.datetime, "%Y-%m-%d")
                    if start_date <= commit_date <= end_date:
                        s_datetime = ' '.join([gitinfo.datetime, gitinfo.time])
                        s = '\t'.join([gitinfo.changeID, gitinfo.commitid,
                                       gitinfo.name, gitinfo.author,
                                       s_datetime, str(gitinfo.change_count),
                                       str(gitinfo.add), str(gitinfo.delete)]
                                      )
                        op.write(s + '\n')
                    # print(gitinfo.commitid, gitinfo.author, gitinfo.name, gitinfo.datetime,
                    # 	  gitinfo.change_count, gitinfo.add, gitinfo.delete, gitinfo.change, sep='\t')


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Please indicate project code dir and output excel file.")
    project_dir = sys.argv[1]
    outputfile = sys.argv[2]
    main(project_dir, outputfile)
