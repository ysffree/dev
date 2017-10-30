import sys
import copy

def handle_repeat(list_all):
    first_list = list(copy.copy(list_all[0]))
    for line in first_list:
        if is_repeat_3(line,list_all):
            first_list.remove(line)
    return first_list




def file2list(vcf):    
    file_list = []
    with open(vcf) as f:
        for line in f:
            line = line.strip('\n')
            if line:
                file_list.append(line)
    return file_list


#def merge(list_all):
#    list_merge = []
#    for i in list_all:
#        list_merge.extend(i)
#    set_merge = set(list_merge)

def be_list(argvs):
    list_all = []
    for vcf in argvs:
        list_all.append(set(file2list(vcf)))
    return list_all
                


def is_repeat_3(line,list_all):
    num = 0
    for vcf in list_all:
        if line in vcf:
            num += 1
    if num > 1:
        return True
    else:
        return False
           


def main():
    argvs = sys.argv[1:]
    list_all = be_list(argvs)
    first_list = handle_repeat(list_all)
    print ('\n'.join(first_list))

if __name__ == '__main__':
    main()

