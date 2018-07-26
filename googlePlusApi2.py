# -*- coding: utf-8 -*-
"""
Created on Wed Jul 25 23:25:01 2018

@author: niles
"""

#import os
import math
import pandas
#import json
from apiclient.discovery import build 
if __name__=='__main__':
   
    coname_twitter_account_file = pandas.read_excel('twitterHandles.xlsx', sheetname='Sheet1') 
    k=0
    for i in coname_twitter_account_file.index:
        k=k+1
        print(k)
        twitterHandleCollections=[]
        coname=coname_twitter_account_file['coname'][i]
        twitterMainHandle=coname_twitter_account_file['twitter main account'][i]        
        twitterHandleCollections.append(coname_twitter_account_file['twitterOtherHandle1'][i])
        twitterHandleCollections.append(coname_twitter_account_file['twitterOtherHandle2'][i])
        twitterHandleCollections.append(coname_twitter_account_file['twitterOtherHandle3'][i])
        twitterHandleCollections.append(coname_twitter_account_file['twitterOtherHandle4'][i])
        twitterHandleCollections.append(coname_twitter_account_file['twitterOtherHandle5'][i])
        twitterHandleCollections.append(coname_twitter_account_file['twitterOtherHandle6'][i])
        twitterHandleCollections.append(coname_twitter_account_file['twitterOtherHandle7'][i])
        twitterHandleCollections.append(coname_twitter_account_file['twitterOtherHandle8'][i])
        twitterHandleCollections.append(coname_twitter_account_file['twitterOtherHandle9'][i])
        twitterUserMainHandle=None
        
        try:
            if pandas.isnull(twitterMainHandle):
                try:
                    twitterMainHandle=None
                except:
                    pass
            else:
                url_part=twitterMainHandle.replace('@','+').lower()
                try:
                    api_key='AIzaSyCQNFFvQ7PEt8TeW7vjJSRN8yk2IXSA9fI'
                    service=build('plus','v1',developerKey=api_key)
                    p_resource=service.people()
                    #service, flags = sample_tools.init(argv, 'plus', 'v1', __doc__, __file__,scope='https://www.googleapis.com/auth/plus.me')
                    pd= p_resource.get(userId=url_part).execute()
                    outputFile=url_part+'.xlsx'
                    writer = pandas.ExcelWriter(outputFile, engine='xlsxwriter')
                    request = service.activities().list(userId=pd['id'], collection='public')
                    #request = p_resource.listByActivity(activityId='+amazon',collection='plusoners',maxResults='2')    
                    count=0
                    itemDict={}
                    repliesArr=[]
                    titleArr=[]
                    publishedDateArr=[]
                    plusonersArr=[]
                    resharersArr=[]
                    while request is not None:
                        activities_doc = request.execute()
                        for item in activities_doc.get('items', []):
                            #print('%-040s -> %s' % (item['id'], item['object']['content'][:30]))
                            itemDict.update({item['id']:item['title']})
                            titleArr.append(item['title'])
                            publishedDateArr.append(item['published'])
                            plusonersArr.append(item['object'].get('plusoners')['totalItems'])
                            repliesArr.append(item['object'].get('replies')['totalItems'])
                            resharersArr.append(item['object'].get('resharers')['totalItems'])
                            count=count+1
                        request = service.activities().list_next(request, activities_doc)
                    df2 = pandas.DataFrame({'Posts': titleArr,'Published Date':publishedDateArr,'PlusOners':plusonersArr,'Replies':repliesArr,'Resharers':resharersArr})
                    df1 = pandas.DataFrame({'Company Name': [pd['displayName']],'Comapany Followers':[pd['circledByCount']],'TagLine':pd['tagline'],'About Me':pd['aboutMe']})                    
                    df1.to_excel(writer, sheet_name='Company Info')
                    df2.to_excel(writer, sheet_name=url_part+'Posts')
                    writer.save()
                    count=0
                    itemDict={}
                    repliesArr=[]
                    titleArr=[]
                    publishedDateArr=[]
                    plusonersArr=[]
                    resharersArr=[]
                except Exception as e:
                    print(e)
                    pass
            for handle in twitterHandleCollections:    
                if not pandas.isnull(handle) or not math.isnan(handle):
                    
                    url_part=handle.replace('@','+').lower()
                    try:
                        api_key='AIzaSyCQNFFvQ7PEt8TeW7vjJSRN8yk2IXSA9fI'
                        service=build('plus','v1',developerKey=api_key)
                        p_resource=service.people()
                        #service, flags = sample_tools.init(argv, 'plus', 'v1', __doc__, __file__,scope='https://www.googleapis.com/auth/plus.me')
                        pd= p_resource.get(userId=url_part).execute()
                        outputFile=url_part+'.xlsx'
                        writer = pandas.ExcelWriter(outputFile, engine='xlsxwriter')
                        request = service.activities().list(userId=pd['id'], collection='public')
                        #request = p_resource.listByActivity(activityId='+amazon',collection='plusoners',maxResults='2')    
                        count=0
                        itemDict={}
                        repliesArr=[]
                        titleArr=[]
                        publishedDateArr=[]
                        plusonersArr=[]
                        resharersArr=[]
                        while request is not None:
                            activities_doc = request.execute()
                            for item in activities_doc.get('items', []):
                                #print('%-040s -> %s' % (item['id'], item['object']['content'][:30]))
                                itemDict.update({item['id']:item['title']})
                                titleArr.append(item['title'])
                                publishedDateArr.append(item['published'])
                                plusonersArr.append(item['object'].get('plusoners')['totalItems'])
                                repliesArr.append(item['object'].get('replies')['totalItems'])
                                resharersArr.append(item['object'].get('resharers')['totalItems'])
                                count=count+1
                            request = service.activities().list_next(request, activities_doc)
                        df2 = pandas.DataFrame({'Posts': titleArr,'Published Date':publishedDateArr,'PlusOners':plusonersArr,'Replies':repliesArr,'Resharers':resharersArr})
                        df1 = pandas.DataFrame({'Company Name': [pd['displayName']],'Comapany Followers':[pd['circledByCount']],'TagLine':pd['tagline'],'About Me':pd['aboutMe']})                    
                        df1.to_excel(writer, sheet_name='Company Info')
                        df2.to_excel(writer, sheet_name=url_part+'Posts')
                        writer.save()
                        count=0
                        itemDict={}
                        repliesArr=[]
                        titleArr=[]
                        publishedDateArr=[]
                        plusonersArr=[]
                        resharersArr=[]
                    except Exception as e:
                        print(e)
                        continue    
        
        except:
            continue


    