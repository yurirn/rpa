#!/usr/bin/env python3
"""
Script para corrigir a pasta dist existente
Copia todos os arquivos necessários para o executável funcionar
"""

import os
import shutil

def fix_dist_folder():
    """Corrige a pasta dist copiando arquivos necessários"""
    print("🔧 Corrigindo pasta dist...")
    
    dist_dir = 'dist'
    if not os.path.exists(dist_dir):
        print("❌ Pasta dist não encontrada")
        return False
    
    # Lista de arquivos necessários
    required_files = [
        'config.json',
        'modules.json'
    ]
    
    # Lista de pastas necessárias
    required_dirs = [
        'src',
        'xml'
    ]
    
    print("📁 Copiando arquivos...")
    
    # Copiar arquivos
    for file in required_files:
        if os.path.exists(file):
            dest_file = os.path.join(dist_dir, file)
            shutil.copy2(file, dest_file)
            print(f"✓ Copiado: {file} -> {dest_file}")
        else:
            print(f"⚠️  Arquivo não encontrado: {file}")
    
    print("\n📁 Copiando pastas...")
    
    # Copiar pastas
    for dir_name in required_dirs:
        if os.path.exists(dir_name):
            dest_dir = os.path.join(dist_dir, dir_name)
            
            # Remover pasta de destino se existir
            if os.path.exists(dest_dir):
                shutil.rmtree(dest_dir)
                print(f"🗑️  Removida pasta existente: {dest_dir}")
            
            # Copiar pasta
            shutil.copytree(dir_name, dest_dir)
            print(f"✓ Copiado: {dir_name}/ -> {dest_dir}/")
        else:
            print(f"⚠️  Pasta não encontrada: {dir_name}")
    
    print("\n✅ Correção concluída!")
    print(f"📂 Pasta dist agora contém todos os arquivos necessários")
    
    # Listar conteúdo da pasta dist
    print("\n📋 Conteúdo da pasta dist:")
    for item in os.listdir(dist_dir):
        item_path = os.path.join(dist_dir, item)
        if os.path.isdir(item_path):
            print(f"📁 {item}/")
        else:
            print(f"📄 {item}")
    
    return True

if __name__ == "__main__":
    fix_dist_folder() 