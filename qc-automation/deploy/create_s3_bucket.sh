#!/bin/bash
# ============================================================
# Create S3 Bucket for QC Dashboard Data
# ============================================================
# Run this once to create the S3 bucket that stores QC data.
# Make sure you have AWS CLI configured with appropriate permissions.
#
# Usage:
#   chmod +x create_s3_bucket.sh
#   ./create_s3_bucket.sh
# ============================================================

BUCKET_NAME="${QC_S3_BUCKET:-fincom-qc-data}"
REGION="${AWS_DEFAULT_REGION:-us-east-1}"

echo "Creating S3 bucket: $BUCKET_NAME in $REGION"

# Create bucket
aws s3api create-bucket \
    --bucket "$BUCKET_NAME" \
    --region "$REGION" \
    --create-bucket-configuration LocationConstraint="$REGION" \
    2>/dev/null || echo "Bucket may already exist"

# Enable versioning (protects against accidental overwrite)
aws s3api put-bucket-versioning \
    --bucket "$BUCKET_NAME" \
    --versioning-configuration Status=Enabled

# Block public access (PII data — must be private!)
aws s3api put-public-access-block \
    --bucket "$BUCKET_NAME" \
    --public-access-block-configuration \
        BlockPublicAcls=true,IgnorePublicAcls=true,BlockPublicPolicy=true,RestrictPublicBuckets=true

# Enable server-side encryption
aws s3api put-bucket-encryption \
    --bucket "$BUCKET_NAME" \
    --server-side-encryption-configuration '{
        "Rules": [{
            "ApplyServerSideEncryptionByDefault": {
                "SSEAlgorithm": "AES256"
            },
            "BucketKeyEnabled": true
        }]
    }'

# Set lifecycle rule (keep 90 days of old versions)
aws s3api put-bucket-lifecycle-configuration \
    --bucket "$BUCKET_NAME" \
    --lifecycle-configuration '{
        "Rules": [{
            "ID": "CleanOldVersions",
            "Status": "Enabled",
            "NoncurrentVersionExpiration": {
                "NoncurrentDays": 90
            },
            "Filter": {"Prefix": ""}
        }]
    }'

echo ""
echo "✅ S3 bucket '$BUCKET_NAME' is ready!"
echo ""
echo "Bucket settings:"
echo "  - Versioning: Enabled"
echo "  - Public access: BLOCKED (private only)"
echo "  - Encryption: AES-256 (server-side)"
echo "  - Old versions auto-delete: After 90 days"
echo ""
echo "Create the data prefix:"
aws s3api put-object --bucket "$BUCKET_NAME" --key "current/" --content-length 0
echo "  ✅ Created: s3://$BUCKET_NAME/current/"
echo ""
echo "Next: Attach an IAM role to your EC2 instance with this policy:"
echo ""
cat << 'POLICY'
{
    "Version": "2012-10-17",
    "Statement": [
        {
            "Effect": "Allow",
            "Action": [
                "s3:GetObject",
                "s3:PutObject",
                "s3:ListBucket",
                "s3:DeleteObject"
            ],
            "Resource": [
                "arn:aws:s3:::fincom-qc-data",
                "arn:aws:s3:::fincom-qc-data/*"
            ]
        }
    ]
}
POLICY
