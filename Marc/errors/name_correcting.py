#coding:utf-8
import os, sys
import fitz
import queue
import threading

def correcting(pdf_path, word):
    doc = fitz.open(pdf_path)
    for page in doc:
        areas = page.searchFor(word)
        print(areas)


if __name__ == '__main__':
    txt = open('names.txt', 'r', encoding='utf-8')
    #with open('names.txt', 'r') as txt:
    line = txt.readlines()
    print(line)
    txt.close()
    # for word in line:
    #     print(word.strip())
    names_correcting_path = '/data/names1/'
    if os.path.exists(names_correcting_path):
        print("{}文件夹已存在".format(names_correcting_path))
    else:
        os.mkdir(names_correcting_path)
        print("{}文件夹已创建".format(names_correcting_path))
    pdf_dir_path = '/data/01/'
    # #print(os.path.join(names_correcting_path, os.path.splitext(os.path.basename('/Volumes/资源管理/mnt/B0A0C8694F3404DC4A80420A087BBBDF1000.pdf'))[0]))
    # threadList = ["Thread-1", "Thread-2", "Thread-3", "Thread-4", "Thread-5"]
    # queueLock = threading.Lock()
    # workQueue = queue.Queue(1200)
    # threads = []
    # threadID = 1
    # file = open('result.txt', 'a+', encoding='utf-8')
    #
    #
    # # 创建新线程
    # for tName in threadList:
    #     thread = myThread(threadID, tName, workQueue)
    #     thread.start()
    #     threads.append(thread)
    #     threadID += 1
    # # 填充队列
    # queueLock.acquire()
    # for files in os.listdir(os.chdir(pdf_dir_path)):
    #     if files.endswith('pdf') or files.endswith('PDF'):
    #         pdf_path = os.path.join(pdf_dir_path, files)
    #         print('pdf_path:', pdf_path)
    #         workQueue.put(pdf_path)
    #         # if workQueue.qsize() >= 1000: break
    # print(workQueue.qsize())
    # queueLock.release()
    #file.close()
    # for t in threads:
    #     t.join()

    # doc = fitz.open('/Users/mac5318/Downloads/B0A0C8694F3404DC4A80420A087BBBDF1000.pdf')
    # doc1 = fitz.open()
    # for page in doc:
    #     #for word in line:
    #         #print(word.strip())
    #         #areas = page.searchFor(word.strip()) #三阶行列式
    #         #if len(areas) > 0:
    #     areas = page.searchFor('三阶行列式', quads=True)
    #     if len(areas) > 0:
    #         page.addSquigglyAnnot(areas)
    #         doc1.insertPDF(doc, from_page=page.number, to_page=page.number)
    #         print(areas, '三阶行列式', page.number)

    #doc1.save('/Users/mac5318/Downloads/B0A0C8694F3404DC4A80420A087BBBDF1000_%s_1.pdf'%('三阶行列式'))
    file = open('result.txt', 'w', encoding='utf-8')
    i = 0
    for files in os.listdir(os.chdir(pdf_dir_path)):
        i = i + 1
        print(i,files)
        if i<58:
            continue
        try:
            if files.endswith('pdf') or files.endswith('PDF'):
                pdf_path = os.path.join(pdf_dir_path, files)
                print('序号：', i, pdf_path)
                doc = fitz.open(pdf_path)
                doc1 = fitz.open()
                errors = []
                for page in doc:
                    for word in line:
                        #try:
                        areas = page.searchFor(word.strip(), quads=True)
                        if len(areas) > 0:
                            page.addSquigglyAnnot(areas)
                            doc1.insertPDF(doc, from_page=page.number, to_page=page.number)
                            print(areas, word.strip(), page.number)
                            errors.append(word.strip() + '_' + str(page.number + 1))
                        #except Exception as e :
                            #print(e)
                         #   continue
                if len(errors) > 0:
                    file.write(os.path.splitext(os.path.basename(pdf_path))[0] + ' ' + ';'.join(errors))
                    file.write('\n')
                print(len(doc1), os.path.join(names_correcting_path, os.path.splitext(os.path.basename(pdf_path))[0] + '_1.pdf'))
                if len(doc1) > 0:
                    doc1.save(os.path.join(names_correcting_path, os.path.splitext(os.path.basename(pdf_path))[0] + '_1.pdf'))
                doc1.close()
                doc.close()

        except Exception as e:
            print(e)
            continue
    file.close()

