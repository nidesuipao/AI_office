import os
from minio import Minio
from minio.error import S3Error


def ensure_bucket(client: Minio, bucket_name: str) -> None:
    if not client.bucket_exists(bucket_name):
        client.make_bucket(bucket_name)


def upload_file_to_minio(
    endpoint: str,
    access_key: str,
    secret_key: str,
    bucket_name: str,
    local_path: str,
    object_name: str,
    secure: bool = False,
) -> str:
    client = Minio(endpoint, access_key=access_key, secret_key=secret_key, secure=secure)
    ensure_bucket(client, bucket_name)
    client.fput_object(bucket_name, object_name, local_path)
    # 生成可用 URL（预签名，默认有效期 7 天）
    url = client.get_presigned_url("GET", bucket_name, object_name)
    return url


if __name__ == "__main__":
    endpoint = os.getenv("MINIO_ENDPOINT", "172.27.0.1:19100")
    access_key = os.getenv("MINIO_ACCESS_KEY", "rustfsadmin")
    secret_key = os.getenv("MINIO_SECRET_KEY", "rustfsadmin")
    bucket = os.getenv("MINIO_BUCKET", "ai-office-test")

    local_file = os.path.join(os.path.dirname(os.path.dirname(__file__)), "output_file", "test_pptx_builder.pptx")
    if not os.path.exists(local_file):
        raise FileNotFoundError(f"文件不存在: {local_file}")

    object_name = os.path.basename(local_file)

    try:
        url = upload_file_to_minio(
            endpoint=endpoint,
            access_key=access_key,
            secret_key=secret_key,
            bucket_name=bucket,
            local_path=local_file,
            object_name=object_name,
            secure=False,
        )
        print(f"上传成功: {url}")
    except S3Error as exc:
        print(f"上传失败: {exc}")
