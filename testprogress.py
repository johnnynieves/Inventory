from tqdm import tqdm
from time import sleep

found = ['abc123', 'def456', 'ghi789', 'jkl012', 'mno345', 'pqr678', 'stu901', 'vwx345', 'yz12345', 'xy12345', 'xy12345', 'xy12345', 'xy12345','abc',]
notfound = ['abc123', 'def456', 'ghi789', 'jkl012', 'mno345', 'pqr678', 'stu901', 'vwx345', 'yz12345', 'xy12345', 'xy12345', 'xy12345', 'xy12345','abc',]
print("Items in File")

# for computer in tqdm(range(len(found),ncols=len(found),ascii="[]",desc='Serial Processed: ')):
    
pbar =  tqdm(range(len(found)),desc='Serials Numbers Processing: ')
p2bar = tqdm(range(len(notfound)),desc='inside cell checking')

for i in range(len(found)+1):
    # print(i)
    sleep(.5)
    

    for p in range(len(notfound)+1):
        # print(p)
        sleep(.05)
        p2bar.update(1)
    
    pbar.update()
p2bar.close()
pbar.close()
print("Done")