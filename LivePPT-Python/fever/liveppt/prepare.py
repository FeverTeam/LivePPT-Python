# -*- coding: UTF-8 -*-

from boto.s3.connection import S3Connection
from boto.s3.key import Key

import config as conf
import boto.sqs
import boto.sns

#准备S3服务
def gen_s3():
    s3_conn = S3Connection(conf.LBW_AWS_ACCESS_KEY, conf.LBW_AWS_SECRET_KEY)
    pptstore_bucket = s3_conn.get_bucket(conf.BUCKET_NAME)
    k = Key(pptstore_bucket)
    return s3_conn, pptstore_bucket, k

#准备S3服务
def gen_sqs():
    sqs_conn = boto.sqs.connect_to_region(conf.TOKYO_REGION,\
                    aws_access_key_id = conf.LBW_AWS_ACCESS_KEY,\
                    aws_secret_access_key = conf.LBW_AWS_SECRET_KEY)                
    q = sqs_conn.create_queue(conf.QUEUE_NAME)
    return sqs_conn, q

#准备SNS服务
def gen_sns():
    sns_conn = boto.sns.connect_to_region(conf.TOKYO_REGION,\
                    aws_access_key_id = conf.LBW_AWS_ACCESS_KEY,\
                    aws_secret_access_key = conf.LBW_AWS_SECRET_KEY)
    return sns_conn