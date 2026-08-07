import sys, os, re
sys.stdout.reconfigure(encoding='utf-8')
d = 'input/Проба пера_txt'
for f in sorted(os.listdir(d)):
    text = open(os.path.join(d, f), encoding='utf-8').read()
    lines = text.split('\n')
    cnt = 0
    print(f'=== {f} ===')
    for line in lines:
        if line.strip():
            print(line.strip())
            cnt += 1
            if cnt >= 20:
                break
    print()
sys.stdout.reconfigure(encoding='utf-8')
d = 'input/Проба пера_txt'
for f in sorted(os.listdir(d)):
    text = open(os.path.join(d, f), encoding='utf-8').read()
    lines = text.split('\n')
    cnt = 0
    print(f'=== {f} ===')
    for line in lines:
        if line.strip():
            print(line.strip())
            cnt += 1
            if cnt >= 20:
                break
    print()

