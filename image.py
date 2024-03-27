import boto3

def convert_jan_to_asin(jan_list):
    # Set up the client for the Product Advertising API
    aws_access_key_id = "YOUR_AWS_ACCESS_KEY_ID"
    aws_secret_access_key = "YOUR_AWS_SECRET_ACCESS_KEY"
    aws_associate_tag = "YOUR_AWS_ASSOCIATE_TAG"
    region_name = "us-west-2" # Change this to your desired region
    client = boto3.client("s3", aws_access_key_id=aws_access_key_id,
                          aws_secret_access_key=aws_secret_access_key,
                          region_name=region_name)
    
    # Make a request for each JAN code and extract the ASIN
    asin_list = []
    for jan in jan_list:
        response = client.item_lookup(ItemId=jan, IdType="EAN", SearchIndex="All",
                                       AssociateTag=aws_associate_tag)
        asin = response["Items"]["Item"]["ASIN"]
        asin_list.append(asin)
    
    return asin_list

# Example usage
jan_list = ["4902505220188", "4901301016354", "4902102111239"]
asin_list = convert_jan_to_asin(jan_list)
print(asin_list)
