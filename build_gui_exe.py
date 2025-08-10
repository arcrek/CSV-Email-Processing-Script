#!/usr/bin/env python3
"""
Optimized build script for GUI version - creates smaller executable
"""

import subprocess
import sys
import os
import shutil

def build_optimized_executable():
    """Build optimized executable with reduced file size"""
    
    print("Building Optimized CSV Email Processor GUI...")
    print("=" * 60)
    
    # Clean previous builds
    if os.path.exists('dist'):
        shutil.rmtree('dist')
        print("Cleaned previous dist folder")
    
    if os.path.exists('build'):
        shutil.rmtree('build')
        print("Cleaned previous build folder")
    
    # Optimized PyInstaller command
    cmd = [
        'pyinstaller',
        '--onefile',                                    # Single executable
        '--windowed',                                   # No console window
        '--optimize=2',                                 # Maximum bytecode optimization
        '--name=CSVEmailProcessorGUI',                  # Executable name
        
        # Exclude only safe-to-remove modules (more conservative approach)
        '--exclude-module=matplotlib',
        '--exclude-module=test',
        '--exclude-module=unittest',
        '--exclude-module=doctest',
        '--exclude-module=pdb',
        '--exclude-module=PIL',
        '--exclude-module=setuptools',
        '--exclude-module=pkg_resources',
        
        # Essential hidden imports for pandas/numpy
        '--hidden-import=secrets',
        '--hidden-import=pandas',
        '--hidden-import=pandas._libs.tslibs.base',
        '--hidden-import=pandas._libs.tslibs.nattype',
        '--hidden-import=pandas._libs.tslibs.np_datetime',
        '--hidden-import=pandas._libs.tslibs.timedeltas',
        '--hidden-import=pandas._libs.tslibs.timestamps',
        '--hidden-import=numpy',
        '--hidden-import=numpy.random',
        '--hidden-import=numpy.random._bounded_integers',
        '--hidden-import=numpy.random.bit_generator',
        '--hidden-import=numpy.core',
        '--hidden-import=numpy.core._multiarray_umath',
        
        # Strip debug symbols
        '--strip',
        
        # Source file
        'csv_email_processor_gui.py'
    ]
    
    try:
        print("Running PyInstaller with optimization...")
        print("This may take several minutes...")
        print("")
        
        # Show the command being run for debugging
        print("PyInstaller command:")
        print(" ".join(cmd))
        print("")
        
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        
        # Get file size
        exe_path = 'dist/CSVEmailProcessorGUI.exe'
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            print(f"\n✓ Build successful!")
            print(f"✓ Executable size: {size_mb:.1f} MB")
            print(f"✓ Location: {exe_path}")
            
            # Try UPX compression if available
            try_upx_compression(exe_path)
            
        else:
            print("❌ Executable not found after build")
            return False
        
        print("\nOptimizations applied:")
        print("  - Bytecode optimization (--optimize=2)")
        print("  - Unnecessary modules excluded")
        print("  - Debug symbols stripped")
        print("  - Windowed mode (no console)")
        
        print(f"\nYou can now run: ./{exe_path}")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Build failed: {e}")
        print(f"PyInstaller output: {e.stdout}")
        print(f"PyInstaller errors: {e.stderr}")
        return False
    
    except FileNotFoundError:
        print("❌ PyInstaller not found. Installing...")
        try:
            subprocess.run([sys.executable, '-m', 'pip', 'install', 'pyinstaller'], check=True)
            print("✓ PyInstaller installed. Please run the build script again.")
        except subprocess.CalledProcessError:
            print("❌ Failed to install PyInstaller")
        return False

def try_upx_compression(exe_path):
    """Try to compress with UPX if available"""
    try:
        # Check if UPX is available
        subprocess.run(['upx', '--version'], check=True, capture_output=True)
        
        # Get original size
        original_size = os.path.getsize(exe_path) / (1024 * 1024)
        
        print("\n🔄 Applying UPX compression...")
        
        # Compress with UPX
        result = subprocess.run(['upx', '--best', '--lzma', exe_path], 
                              check=True, capture_output=True, text=True)
        
        # Get compressed size
        compressed_size = os.path.getsize(exe_path) / (1024 * 1024)
        savings = ((original_size - compressed_size) / original_size) * 100
        
        print(f"✓ UPX compression successful!")
        print(f"  Original: {original_size:.1f} MB")
        print(f"  Compressed: {compressed_size:.1f} MB")
        print(f"  Savings: {savings:.1f}%")
        
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("\nℹ️  UPX not available (optional compression tool)")
        print("   Install UPX for additional file size reduction")
        print("   Download from: https://upx.github.io/")

def main():
    """Main build function"""
    
    # Check if source file exists
    if not os.path.exists('csv_email_processor_gui.py'):
        print("❌ csv_email_processor_gui.py not found!")
        print("   Please run this script from the same directory as the GUI file.")
        return
    
    # Check Python version
    if sys.version_info < (3, 6):
        print("❌ Python 3.6 or higher required")
        return
    
    print(f"Python version: {sys.version}")
    print(f"Current directory: {os.getcwd()}")
    
    # Check if pandas is available
    try:
        import pandas
        print(f"Pandas version: {pandas.__version__}")
    except ImportError:
        print("❌ Pandas not found. Installing...")
        try:
            subprocess.run([sys.executable, '-m', 'pip', 'install', 'pandas'], check=True)
            print("✓ Pandas installed successfully")
        except subprocess.CalledProcessError:
            print("❌ Failed to install pandas")
            return
    
    # Build the executable
    success = build_optimized_executable()
    
    if success:
        print("\n" + "=" * 60)
        print("🎉 BUILD COMPLETED SUCCESSFULLY!")
        print("\nFiles created:")
        print("  📁 dist/CSVEmailProcessorGUI.exe (main executable)")
        print("  📁 build/ (temporary build files - can be deleted)")
        print("\nThe executable is optimized for minimal size and runs without")
        print("showing a console window. Perfect for distribution!")
    else:
        print("\n❌ Build failed. Please check the error messages above.")

if __name__ == "__main__":
    main()
