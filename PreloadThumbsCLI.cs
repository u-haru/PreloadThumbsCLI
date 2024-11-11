using System.Runtime.InteropServices;
using System.Drawing;
using ShellProgressBar;

namespace PreloadThumbsCLI
{
	class Program
	{
		static Guid shellItemGuid = new Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE");
		static void Main(string[] args)
		{
			if (args.Length == 0)
			{
				Console.WriteLine("Usage: PreloadThumbsCLI <folder1> <folder2> ...");
				return;
			}

			// すべてのフォルダからファイル一覧を作成
            List<string> files = new List<string>();
            foreach (var folderPath in args)
            {
                if (Directory.Exists(folderPath))
                {
                    files.AddRange(Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories));
                }
                else
                {
                    Console.WriteLine($"Folder not found: {folderPath}");
                }
            }
			if (files.Count == 0)
			{
				Console.WriteLine("No files found to process.");
				return;
			}

			// プログレスバーの設定
			var options = new ProgressBarOptions
			{
				ForegroundColor = ConsoleColor.Cyan,
				BackgroundColor = ConsoleColor.DarkGray,
				ProgressCharacter = '─',
				CollapseWhenFinished = true
			};

			// 並列処理でキャッシュを実行
			using (var progressBar = new ProgressBar(files.Count, "Caching images...", options))
			{
				Parallel.ForEach(files, filePath =>
				{
					CacheImage(filePath);

					// プログレスバーの更新をスレッドセーフに行う
					lock (progressBar)
					{
						progressBar.Tick($"Processed: {filePath}");
					}
				});
			}
		}

		static void CacheImage(string filePath)
		{
			try
			{
				string absolutePath = Path.GetFullPath(filePath);

				var shellItemImageFactory = (IShellItemImageFactory)CreateShellItem(absolutePath);

				Size size = new Size(256, 256);
				IntPtr hBitmap;
				shellItemImageFactory.GetImage(size, 0x0, out hBitmap);

				// if (hBitmap != IntPtr.Zero)	Console.WriteLine($"Image cached for file: {filePath}");
			}
			catch (Exception ex)
			{
				Console.WriteLine($"Error caching image for file {filePath}: {ex.Message}");
			}
		}

		private static IShellItem CreateShellItem(string filePath)
		{
			IShellItem shellItem;
			int hr = SHCreateItemFromParsingName(filePath, IntPtr.Zero, ref shellItemGuid, out shellItem);
			if (hr != 0) Marshal.ThrowExceptionForHR(hr);
			return shellItem;
		}

		[DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
		private static extern int SHCreateItemFromParsingName(
			[MarshalAs(UnmanagedType.LPWStr)] string pszPath,
			IntPtr pbc,
			ref Guid riid,
			[MarshalAs(UnmanagedType.Interface)] out IShellItem ppv
		);
	}

	[ComImport]
	[Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	interface IShellItem
	{
	}

	[ComImport]
	[Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b")]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	interface IShellItemImageFactory
	{
		[PreserveSig]
		int GetImage(Size size, int flags, out IntPtr phbm);
	}
}
