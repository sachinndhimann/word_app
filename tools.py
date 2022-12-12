import boto3
import io
import docx

s3 = boto3.client(
   "s3",
   aws_access_key_id='',
   aws_secret_access_key=''
)


def upload_file_to_s3(file, bucket_name, acl="public-read"):

    try:

        s3.upload_fileobj(
            file,
            bucket_name,
            file.filename,
            ExtraArgs={
                "ACL": acl,
                "ContentType": file.content_type
            }
        )

    except Exception as e:
        print("Something Happened: ", e)
        return e

    return "{} uploaded to {} ".format(file.filename, bucket_name)

def read_object(bucket,filename):
    bucket = s3.Bucket(bucket)
    object_in_s3 = bucket.Object(filename)
    object_as_streaming_body = object_in_s3.get()["Body"]
    print(f"Type of object_as_streaming_body: {type(object_as_streaming_body)}")
    object_as_bytes = object_as_streaming_body.read()
    print(f"Type of object_as_bytes: {type(object_as_bytes)}")

    # Now we use BytesIO to create a file-like object from our byte-stream
    object_as_file_like = io.BytesIO(object_as_bytes)
        # Et voila!
    document = docx.Document(docx=object_as_file_like)
