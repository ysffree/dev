#!/usr/bin/env python
# coding: utf-8

# **********************************************************************
# file: transvar.py
#
# tools for query transvar web app to trans annotation
# **********************************************************************
"""transvar.py get annotation translate info ({version} by {author})

Usage: transvar.py [opt] gene:mut
   opt:
      -f or --file  string    input is filename
      -s or --save            save each gene result in one file
      -d or --dbset string    query db set, default: refseq, valid db are: {dbset}
      -r or --ref   string    query genome version, default: hg19, valid ref are: {refset}
      -a or --all             query all dbset

example:

transvar.py -f mut.list1 -f mut.list2
transvar.py "met:c.3028G>A" "met:c.3028+1G>T"
transvar.py -d refseq -d ucsc "chr11:g.46761055G>A"
transvar.py -s -f mut.list1

mut list file accept format:
1. GENE mut
   EGFR p.A763_Y764insFQEA

2. GENE:mut
   EGFR:p.A763_Y764insFQEA
"""


from __future__ import print_function
import urllib
import urllib2
import urlparse
import sys
import time
import getopt
import os


__VERSION__ = "2016.09.02.2"
__AUTHOR__ = "azer.xu@gmail.com"

POST_URL = "http://bioinformatics.mdanderson.org/DynamicViewer/"

TASK = (
    "panno", # Reverse Annotation: Protein
    "canno", # Reverse Annotation: cDNA
    "ganno", # Forward Annotation: gDNA
    "codonsearch", # Codon Search: Protein
)

DB_SET = ("ensembl", "ccds", "refseq", "gencode", "ucsc", "aceview")

REF_VERSION = ("hg19", "hg38", "hg18")


# core fuction
def post(query, dbset=None, task="panno", ref="hg19"):
    if not dbset:
        dbset = ["refseq"]
    dic = {
        "app": "transvar",
        "command_id": "transvar",
        "task": task,
        "refversion": ref,
        "typed_identifiers": query,
    }
    for db in dbset:
        dic["check" + db] = db

    print ("Query:", query, file=sys.stderr)
    data = urllib.urlencode(dic)
    _post = urllib2.urlopen(POST_URL, data, timeout=10)
    ret = _post.read()
    code = _post.code
    if code != 200:
        while True:
            res_url = urllib.urlopen(_post.headers['Location']).read().strip()
            if res_url.startswith("requestKey:"):
                print (res_url, file=sys.stderr)
                time.sleep(1)
                continue
            break
    else:
        res_url = ret.strip()

    url = urlparse.urljoin(res_url.strip().strip("output"), "stdio/stdout.txt")
    return urllib.urlopen(url).read()


def xopen(filename):
    if filename == "-":
        return sys.stdin
    return open(filename)


# accept mode
# GENE mut (example: EGFR p.A763_Y764insFQEA)
# GENE:mut (example: EGFR:p.A763_Y764insFQEA)
def load_file(filename):
    with xopen(filename) as handle:
        for line in handle:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            yield line


def check_query(query):
    try:
        gene, mut = query.split(":", 1) if ":" in query else query.split(None, 1)
    except:
        print (query)
        raise

    if mut.startswith("p."):
        return "panno", "%s:%s" % (gene.upper(), mut.rstrip("."))
    elif mut.startswith("c."):
        return "canno", "%s:%s" % (gene.upper(), check_mut(mut))
    elif mut.startswith("g."):
        #gene = gene.lower()
        if 'del' not in mut and 'ins' not in mut:
            mut = parse_mut(mut)
        if not gene.startswith("chr"):
            gene = "chr" + gene
        return "ganno", "%s:%s" % (gene, check_mut(mut))

    print ("unkown type:", gene, mut, file=sys.stderr)
    return "", ""

def parse_mut(mut):
    mut = mut.strip('\n')
    if not mut[3].isdigit() or not mut[-2].isdigit():
        list_all = change_mut(mut)
        num_f = str(int(list_all[1]) + len(list_all[0])-1)
        mut = '{0}_{1}del{2}ins{3}'.format(list_all[1],num_f,list_all[0],list_all[2])
    return mut

def change_mut(mut):
    mut = mut.split(".", 1)[1]
    mut_list = list(mut)
    del_list = []
    num_list = []
    for i in mut_list:
        if not i.isdigit():
            del_list.append(i)
        else:
            break
    for j in del_list:
        mut_list.remove(j)
    for k in mut_list:
        if k.isdigit():
            num_list.append(k)
        else:
            break
    for m in num_list:
        mut_list.remove(m)
    dl = "".join(del_list)
    num = "".join(num_list)
    ins = "".join(mut_list)
    return [dl, num, ins]

def parse_change(change):
    pos = []
    for b in change.upper():
        if b in 'ATCGatcg':
            if pos:
                yield "".join(pos)
            yield b
        else:
            pos.append(b)


def check_mut(mut):
    if len(mut) > 2 and mut[2].isdigit():
        return mut
    _type, change = mut.split(".")
    g, pos, m = parse_change(change)
    return "%s.%s%s>%s" % (_type, pos, g, m)


def parse_transvar(result):
    for line in result.split("\n"):
        line = line.strip()
        if not line or line.startswith("input"):
            continue
        if "warning:" in line:
            continue
        try:
            query, transcript, gene, strand, coor, region, info = line.split("\t")
        except:
            print (line, file=sys.stderr)
            raise
        transcript = transcript.split(None, 1)[0]
        gDNA, cDNA, protein = coor.split("/")
        yield query, transcript, gene, strand, gDNA, cDNA, protein, region, info


def get_transvar_result(result):
    res = list(parse_transvar(result))
    return ["\t".join(r) for r in res]
    if len(res) == 0:
        return None
    if len(res) == 1:
        return "\t".join(res[0])

    for query, transcript, gene, strand, gDNA, cDNA, protein, region, info in res:
        if not transcript.startswith("NM_"):
            continue
        return "\t".join((query, transcript, gene, strand, gDNA, cDNA, protein, region, info))


def show_usage():
    print (__doc__.format(version=__VERSION__,
                          author=__AUTHOR__,
                          dbset=",".join(DB_SET),
                          refset=",".join(REF_VERSION)), file=sys.stderr)
    exit(1)


def get_opt():
    is_save = False
    input_files = []
    dbset = set()
    ref = "hg19"
    try:
        optlist, args = getopt.getopt(sys.argv[1:], "r:d:f:sa", ["ref=", "db=", "file=", "save", "all"])

        for opt, val in optlist:
            if opt in ("-f", "--file"):
                input_files.append(val)
            elif opt in ("-r", "--ref"):
                val = val.lower()
                if val not in REF_VERSION:
                    print (val, "not in valid ref set:", ",".join(REF_VERSION))
                    exit(1)
                ref = val
            elif opt in ("-d", "--db"):
                val = val.lower()
                if val not in DB_SET:
                    print (val, "not in valid db set:", ",".join(DB_SET))
                    exit(1)
                dbset.add(val)
            elif opt in ("-s", "--save"):
                is_save = True
            elif opt in ("-a", "--all"):
                dbset = set(DB_SET)
            else:
                show_usage()
    except getopt.GetoptError as e:
        print (e, file=sys.stderr)
        show_usage()

    if not args and not input_files:
        show_usage()

    if not dbset:
        dbset = ["refseq"]
    return input_files, is_save, dbset, ref, args


def load_reg(filename=".transvar.reg"):
    reg = set()

    if not os.path.exists(filename):
        return reg

    with open(filename) as handle:
        for line in handle:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            reg.add(line)
    return reg


def load_bad_reg(filename=".transvar.bad.reg"):
    reg = set()

    if not os.path.exists(filename):
        return reg

    with open(filename) as handle:
        for line in handle:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            reg.add(line)
    return reg


def save_reg(reg, filename=".transvar.reg"):
    with open(filename, "w") as out:
        for q in reg:
            print (q, file=out)


def save_bad_reg(reg, filename=".transvar.bad.reg"):
    with open(filename, "w") as out:
        for q in reg:
            print (q, file=out)


def load_query(querys, input_files):
    for query in querys:
        yield query

    for ifile in input_files:
        for query in load_file(ifile):
            yield query


def main():
    input_files, is_save, dbset, ref, querys = get_opt()

    if is_save:
        reg = load_reg()
        bad_reg = load_bad_reg()

    try:
        for query in load_query(querys, input_files):
            items = query.split(None)
            if len(items) > 1:
                query = ":".join(items)
            try:
                task, q = check_query(query)
                if not task:
                    print ("Error no task:", query, file=sys.stderr)
                    continue

                if not is_save:
                    print (post(q, dbset=dbset, ref=ref, task=task))
                    continue

                if q in reg or q in bad_reg:
                    continue

                ret = post(q, dbset=dbset, ref=ref, task=task)
                gene = query.split(":", 1)[0].upper()
                with  open(gene + ".transvar", "a") as out:
                    res = get_transvar_result(ret)
                    if len(res) == 0:
                        print ("query", q, "no result get!", file=sys.stderr)
                    else:
                        for r in res:
                            print (r, file=out)
                reg.add(q)
            except ValueError as e:
                print ("EE:", e, file=sys.stderr)
                bad_reg.add(q)
                continue
            except IOError:
                print ("IOError:", q, file=sys.stderr)
                time.sleep(10)
                continue
            except KeyboardInterrupt:
                print ("Cancelled by user !!", file=sys.stderr)
                break
    finally:
        if is_save and reg:
            save_reg(reg)
            save_bad_reg(bad_reg)


def test():
    QUERY = "\n".join(("chr11:g.46761055G>A", "chr11:g.46761055G>A"))
    QUERY = "MET:c.3028+1G>T"
    QUERY = "TSC1:p.W103*"
    QUERY = "TSC1:p.W103"
    print (post(QUERY, task="panno"))


if __name__ == "__main__":
    main()
