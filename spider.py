#!/usr/bin/env python
# -*- coding: utf-8 -*-

import requests
import openpyxl
from datetime import datetime, timedelta, timezone
import sys
from getpass import getpass


def setup_proxy(username, password, http_proxy, https_proxy):
    """设置代理配置"""
    if username and password:
        # 在代理URL中添加认证信息
        http_proxy_with_auth = http_proxy.replace('http://', f'http://{username}:{password}@')
        https_proxy_with_auth = https_proxy.replace('http://', f'http://{username}:{password}@')
        proxies = {
            'http': http_proxy_with_auth,
            'https': https_proxy_with_auth
        }
    else:
        proxies = {
            'http': http_proxy,
            'https': https_proxy
        }
    return proxies


def read_npm_packages(excel_file):
    """从Excel文件读取npm库名称列表"""
    try:
        wb = openpyxl.load_workbook(excel_file)
        ws = wb.active
        packages = []
        
        # 读取第一列的所有非空值（跳过表头）
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0]:  # 如果第一列有值
                packages.append(str(row[0]).strip())
        
        wb.close()
        return packages
    except Exception as e:
        print(f"读取Excel文件失败: {e}")
        sys.exit(1)


def get_package_versions(package_name, access_token, proxies):
    """获取npm包的版本信息"""
    url = f"https://registry.npmjs.org/{package_name}"
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json'
    }
    
    try:
        response = requests.get(url, headers=headers, proxies=proxies, timeout=30)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"获取包 {package_name} 信息失败: {e}")
        return None


def filter_versions_last_year(package_data):
    """筛选最近一年的版本信息"""
    if not package_data or 'time' not in package_data:
        return []
    
    one_year_ago = datetime.now(timezone.utc) - timedelta(days=365)
    versions_info = []
    
    time_info = package_data.get('time', {})
    versions = package_data.get('versions', {})
    
    for version, publish_time in time_info.items():
        if version in ['created', 'modified']:
            continue
        
        try:
            publish_date = datetime.fromisoformat(publish_time.replace('Z', '+00:00'))
            if publish_date >= one_year_ago:
                version_data = versions.get(version, {})
                versions_info.append({
                    'version': version,
                    'publish_time': publish_time,
                    'description': version_data.get('description', ''),
                    'author': version_data.get('author', {}).get('name', '') if isinstance(version_data.get('author'), dict) else str(version_data.get('author', '')),
                    'dependencies': len(version_data.get('dependencies', {}))
                })
        except Exception as e:
            print(f"处理版本 {version} 时出错: {e}")
            continue
    
    # 按发布时间排序
    versions_info.sort(key=lambda x: x['publish_time'], reverse=True)
    return versions_info


def write_results_to_excel(results, output_file):
    """将扫描结果写入Excel文件"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "扫描结果"
    
    # 写入表头
    headers = ['包名', '版本', '发布时间', '描述', '作者', '依赖数量']
    ws.append(headers)
    
    # 写入数据
    for package_name, versions in results.items():
        if not versions:
            ws.append([package_name, '未找到最近一年的版本', '', '', '', ''])
        else:
            for version_info in versions:
                ws.append([
                    package_name,
                    version_info['version'],
                    version_info['publish_time'],
                    version_info['description'],
                    version_info['author'],
                    version_info['dependencies']
                ])
    
    # 调整列宽
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width
    
    wb.save(output_file)
    print(f"\n扫描结果已保存到: {output_file}")


def main():
    print("=" * 60)
    print("NPM库版本扫描工具")
    print("=" * 60)
    
    # 1. 获取代理配置
    print("\n请输入代理配置（如果不需要代理，直接按回车跳过）:")
    proxy_username = input("代理用户名: ").strip()
    proxy_password = ""
    if proxy_username:
        proxy_password = getpass("代理密码: ")
    
    http_proxy = input("HTTP Proxy (例如: http://proxy.huawei.com:8080): ").strip()
    https_proxy = input("HTTPS Proxy (例如: http://proxy.huawei.com:8080): ").strip()
    
    proxies = None
    if http_proxy or https_proxy:
        proxies = setup_proxy(proxy_username, proxy_password, http_proxy, https_proxy)
        print("✓ 代理配置完成")
    
    # 2. 获取输入文件路径
    print("\n请输入包含npm库名称的Excel文件路径:")
    input_file = input("Excel文件路径: ").strip()
    if not input_file:
        print("错误: 必须提供Excel文件路径")
        sys.exit(1)
    
    # 3. 获取npm access token
    print("\n请输入npm access token:")
    access_token = getpass("Access Token: ").strip()
    if not access_token:
        print("错误: 必须提供npm access token")
        sys.exit(1)
    
    # 4. 读取npm包列表
    print(f"\n正在读取Excel文件: {input_file}")
    packages = read_npm_packages(input_file)
    print(f"✓ 共读取到 {len(packages)} 个npm包")
    
    # 5. 扫描每个包的版本信息
    print("\n开始扫描npm包版本信息...")
    results = {}
    
    for i, package in enumerate(packages, 1):
        print(f"[{i}/{len(packages)}] 正在扫描: {package}")
        package_data = get_package_versions(package, access_token, proxies)
        
        if package_data:
            versions = filter_versions_last_year(package_data)
            results[package] = versions
            print(f"  ✓ 找到 {len(versions)} 个最近一年的版本")
        else:
            results[package] = []
            print(f"  ✗ 获取失败")
    
    # 6. 输出结果到Excel
    print("\n正在生成结果文件...")
    output_file = input_file.replace('.xlsx', '-扫描结果.xlsx')
    if output_file == input_file:
        output_file = input_file.replace('.xlsx', '') + '-扫描结果.xlsx'
    
    write_results_to_excel(results, output_file)
    
    # 统计信息
    total_versions = sum(len(versions) for versions in results.values())
    print("\n" + "=" * 60)
    print("扫描完成!")
    print(f"共扫描 {len(packages)} 个npm包")
    print(f"找到 {total_versions} 个最近一年的版本")
    print("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n程序已被用户中断")
        sys.exit(0)
    except Exception as e:
        print(f"\n程序执行出错: {e}")
        sys.exit(1)
