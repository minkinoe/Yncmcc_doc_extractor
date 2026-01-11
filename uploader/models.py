from django.db import models
from django.utils import timezone
import os


def file_upload_path(instance, filename):
    """为上传的文件生成存储路径"""
    # 使用上传时间和原始文件名构建路径
    upload_time = timezone.now().strftime('%Y/%m/%d')
    return os.path.join('uploads', upload_time, filename)


class UploadedFile(models.Model):
    """上传文件模型"""
    # 文件信息
    file = models.FileField(upload_to=file_upload_path, null=True, blank=True)
    original_filename = models.CharField(max_length=255)
    file_size = models.IntegerField(help_text="文件大小（字节）", default=0)
    file_type = models.CharField(max_length=50, help_text="文件类型，如zip、doc、docx")
    group_name = models.CharField(max_length=255, null=True, blank=True, help_text="集团名称（来自ZIP文件名）")
    address = models.CharField(max_length=255, null=True, blank=True, help_text="地址（来自ZIP文件名）")
    township = models.CharField(max_length=255, null=True, blank=True, help_text="街道（高德解析，township）")
    construction_unit = models.CharField(max_length=255, null=True, blank=True, help_text="施工单位（由街道映射）")
    is_marked = models.BooleanField(default=True, help_text="是否标记（标星）")
    
    # 处理信息
    uploaded_at = models.DateTimeField(auto_now_add=True)
    processed_at = models.DateTimeField(null=True, blank=True)
    is_processed = models.BooleanField(default=False)
    processing_error = models.TextField(null=True, blank=True)
    
    # 统计信息
    document_count = models.IntegerField(default=0, help_text="提取到的文档数量")
    
    def __str__(self):
        return self.original_filename
    
    class Meta:
        ordering = ['-uploaded_at']


class ExtractedInfo(models.Model):
    """从文档中提取的信息模型"""
    # 关联到上传的文件
    uploaded_file = models.ForeignKey(UploadedFile, on_delete=models.CASCADE, related_name='extracted_infos')
    
    # 文档基本信息
    order_code = models.CharField(max_length=100, help_text="单号", null=True, blank=True, db_index=True)
    document_name = models.CharField(max_length=255, help_text="文档文件名", default="")
    document_content = models.TextField(help_text="从Word中提取的完整文本", null=True, blank=True)
    extraction_status = models.CharField(max_length=20, default="待处理")
    extraction_error = models.TextField(null=True, blank=True)
    
    # 提取的价格信息
    maintenance_fee = models.DecimalField(max_digits=10, decimal_places=2, default=0.00, help_text="维护费")
    service_fee = models.DecimalField(max_digits=10, decimal_places=2, default=0.00, help_text="服务费")
    terminal_fee = models.DecimalField(max_digits=10, decimal_places=2, default=0.00, help_text="终端费")
    other_fees = models.DecimalField(max_digits=10, decimal_places=2, default=0.00, help_text="其他费用")
    total_fees = models.DecimalField(max_digits=10, decimal_places=2, default=0.00, help_text="费用总计")
    doc_maintenance_total = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True, help_text="文档中维护费合计")
    overall_total_price = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True, help_text="总体花费")
    total_price = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True, help_text="总花费")
    
    # 光缆信息 - 现在支持多条光缆记录
    fiber_info = models.JSONField(null=True, blank=True, help_text="光缆信息列表")
    
    # 设备信息
    equipment_items = models.JSONField(null=True, blank=True, help_text="设备清单")
    
    # 验证信息
    verification_passed = models.BooleanField(default=False)
    verification_message = models.TextField(null=True, blank=True, help_text="验证消息")
    
    # 时间信息
    extracted_at = models.DateTimeField(auto_now_add=True)
    
    def __str__(self):
        return f"{self.order_code} - {self.document_name}"
    
    class Meta:
        ordering = ['-extracted_at']
        indexes = [
            models.Index(fields=['order_code']),
            models.Index(fields=['uploaded_file']),
        ]
