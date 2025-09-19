#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
MinIO文件服务 - 基于MinIO的文件上传和管理服务

替换原有的RustFSService，提供文件上传、下载和管理功能
"""

import os
from typing import Optional
from datetime import timedelta
from minio import Minio
from minio.error import S3Error


class MinIOService:
    """MinIO文件服务类，提供文件上传、下载和管理功能"""
    
    def __init__(self, 
                 endpoint: Optional[str] = None,
                 access_key: Optional[str] = None, 
                 secret_key: Optional[str] = None,
                 bucket_name: Optional[str] = None,
                 secure: bool = False):
        """
        初始化MinIO服务
        
        Args:
            endpoint: MinIO服务端点，默认从环境变量MINIO_ENDPOINT获取
            access_key: 访问密钥，默认从环境变量MINIO_ACCESS_KEY获取
            secret_key: 秘密密钥，默认从环境变量MINIO_SECRET_KEY获取
            bucket_name: 存储桶名称，默认从环境变量MINIO_BUCKET获取
            secure: 是否使用HTTPS，默认False
        """
        self.endpoint = endpoint or os.getenv("MINIO_ENDPOINT", "172.27.0.1:19100")
        self.access_key = access_key or os.getenv("MINIO_ACCESS_KEY", "rustfsadmin")
        self.secret_key = secret_key or os.getenv("MINIO_SECRET_KEY", "rustfsadmin")
        self.bucket_name = bucket_name or os.getenv("MINIO_BUCKET", "ai-office-test")
        self.secure = secure
        
        # 初始化MinIO客户端
        self.client = Minio(
            self.endpoint,
            access_key=self.access_key,
            secret_key=self.secret_key,
            secure=self.secure
        )
        
        # 确保存储桶存在
        self._ensure_bucket_exists()
    
    def _ensure_bucket_exists(self):
        """确保存储桶存在，如果不存在则创建"""
        try:
            if not self.client.bucket_exists(self.bucket_name):
                self.client.make_bucket(self.bucket_name)
                print(f"✅ 创建存储桶: {self.bucket_name}")
            else:
                print(f"✅ 存储桶已存在: {self.bucket_name}")
        except S3Error as e:
            print(f"❌ 存储桶操作失败: {e}")
            raise
    
    def upload_file(self, local_path: str, object_name: Optional[str] = None) -> str:
        """
        上传文件到MinIO
        
        Args:
            local_path: 本地文件路径
            object_name: 对象名称，如果不指定则使用文件名
            
        Returns:
            str: 预签名URL
        """
        if not os.path.exists(local_path):
            raise FileNotFoundError(f"文件不存在: {local_path}")
        
        if object_name is None:
            object_name = os.path.basename(local_path)
        
        try:
            # 上传文件
            self.client.fput_object(self.bucket_name, object_name, local_path)
            print(f"✅ 文件上传成功: {object_name}")
            
            # 生成预签名URL
            url = self.client.get_presigned_url("GET", self.bucket_name, object_name)
            return url
            
        except S3Error as e:
            print(f"❌ 文件上传失败: {e}")
            raise
    
    def download_file(self, object_name: str, local_path: str) -> bool:
        """
        从MinIO下载文件
        
        Args:
            object_name: 对象名称
            local_path: 本地保存路径
            
        Returns:
            bool: 下载是否成功
        """
        try:
            # 确保本地目录存在
            os.makedirs(os.path.dirname(local_path), exist_ok=True)
            
            # 下载文件
            self.client.fget_object(self.bucket_name, object_name, local_path)
            print(f"✅ 文件下载成功: {object_name} -> {local_path}")
            return True
            
        except S3Error as e:
            print(f"❌ 文件下载失败: {e}")
            return False
    
    def delete_file(self, object_name: str) -> bool:
        """
        删除MinIO中的文件
        
        Args:
            object_name: 对象名称
            
        Returns:
            bool: 删除是否成功
        """
        try:
            self.client.remove_object(self.bucket_name, object_name)
            print(f"✅ 文件删除成功: {object_name}")
            return True
            
        except S3Error as e:
            print(f"❌ 文件删除失败: {e}")
            return False
    
    def list_files(self, prefix: str = "") -> list:
        """
        列出存储桶中的文件
        
        Args:
            prefix: 文件名前缀过滤
            
        Returns:
            list: 文件对象列表
        """
        try:
            objects = self.client.list_objects(self.bucket_name, prefix=prefix)
            return [obj.object_name for obj in objects]
            
        except S3Error as e:
            print(f"❌ 文件列表获取失败: {e}")
            return []
    
    def get_file_url(self, object_name: str, expires_in_seconds: int = 604800) -> str:
        """
        获取文件的预签名URL
        
        Args:
            object_name: 对象名称
            expires_in_seconds: URL过期时间（秒），默认7天
            
        Returns:
            str: 预签名URL
        """
        try:
            url = self.client.get_presigned_url(
                "GET", 
                self.bucket_name, 
                object_name,
                expires=timedelta(seconds=expires_in_seconds)
            )
            return url
            
        except S3Error as e:
            print(f"❌ URL生成失败: {e}")
            raise
    
    def file_exists(self, object_name: str) -> bool:
        """
        检查文件是否存在
        
        Args:
            object_name: 对象名称
            
        Returns:
            bool: 文件是否存在
        """
        try:
            self.client.stat_object(self.bucket_name, object_name)
            return True
        except S3Error:
            return False
    
    def get_base_url(self) -> str:
        """
        获取MinIO服务的基础URL
        
        Returns:
            str: 基础URL
        """
        protocol = "https" if self.secure else "http"
        return f"{protocol}://{self.endpoint}"
    
    def get_console_url(self) -> str:
        """
        获取MinIO控制台URL
        
        Returns:
            str: 控制台URL
        """
        protocol = "https" if self.secure else "http"
        # 假设控制台端口是API端口+1
        console_endpoint = self.endpoint.replace(":19100", ":19101")
        return f"{protocol}://{console_endpoint}"
    
    def get_bucket_url(self) -> str:
        """
        获取存储桶的URL
        
        Returns:
            str: 存储桶URL
        """
        base_url = self.get_base_url()
        return f"{base_url}/{self.bucket_name}"


# 为了兼容性，创建一个别名
RustFSService = MinIOService
