using System.Runtime.InteropServices;
using System.Drawing;
using ShellProgressBar;

// 参考: https://github.com/bruhov/WinThumbsPreloader
namespace PreloadThumbsCLI
{
	class Program
	{
		static Guid shellItemGuid = new Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE");
		static Guid CLSIDLocalThumbnailCache = new Guid("50ef4544-ac9f-4a8e-b21b-8a26180db13f");
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
			ISharedBitmap bmp = null;
			WTS_CACHEFLAGS cFlags;
			WTS_THUMBNAILID bmpId;
			IShellItem shellItem = null;
			try
			{
				var TBCacheType = Type.GetTypeFromCLSID(CLSIDLocalThumbnailCache);
				var TBCache = (IThumbnailCache)Activator.CreateInstance(TBCacheType);
				string absolutePath = Path.GetFullPath(filePath);

				shellItem = CreateShellItem(absolutePath);
				TBCache.GetThumbnail(shellItem, 128, WTS_FLAGS.WTS_EXTRACTINPROC, out bmp, out cFlags, out bmpId);
			}
			catch (Exception ex)
			{
				Console.WriteLine($"Error caching image for file {filePath}: {ex.Message}");
			}
			if (bmp != null) Marshal.ReleaseComObject(bmp);
			if (shellItem != null) Marshal.ReleaseComObject(shellItem);
			bmp = null;
			shellItem = null;
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

	[StructLayout(LayoutKind.Sequential, Size = 16), Serializable]
	struct WTS_THUMBNAILID
	{
		[MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst = 16)]
		byte[] rgbKey;
	}

	[ComImportAttribute()]
	[GuidAttribute("091162a4-bc96-411f-aae8-c5122cd03363")]
	[InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
	public interface ISharedBitmap
	{
	}

	[Flags]
	enum WTS_FLAGS : uint
	{
		WTS_EXTRACT = 0x00000000,
		WTS_INCACHEONLY = 0x00000001,
		WTS_FASTEXTRACT = 0x00000002,
		WTS_SLOWRECLAIM = 0x00000004,
		WTS_FORCEEXTRACTION = 0x00000008,
		WTS_EXTRACTDONOTCACHE = 0x00000020,
		WTS_SCALETOREQUESTEDSIZE = 0x00000040,
		WTS_SKIPFASTEXTRACT = 0x00000080,
		WTS_EXTRACTINPROC = 0x00000100
	}

	[ComImport]
	[Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	interface IShellItem
	{
	}

	[Flags]
	enum WTS_CACHEFLAGS : uint
	{
		WTS_DEFAULT = 0x00000000,
		WTS_LOWQUALITY = 0x00000001,
		WTS_CACHED = 0x00000002
	}

	[ComImport]
	[Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b")]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	interface IShellItemImageFactory
	{
		[PreserveSig]
		int GetImage(Size size, int flags, out IntPtr phbm);
	}
	[ComImport]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("F676C15D-596A-4ce2-8234-33996F445DB1")]
	interface IThumbnailCache
	{
		uint GetThumbnail(
			[In] IShellItem pShellItem,
			[In] uint cxyRequestedThumbSize,
			[In] WTS_FLAGS flags /*default:  WTS_FLAGS.WTS_EXTRACT*/,
			[Out][MarshalAs(UnmanagedType.Interface)] out ISharedBitmap ppvThumb,
			[Out] out WTS_CACHEFLAGS pOutFlags,
			[Out] out WTS_THUMBNAILID pThumbnailID
		);

		void GetThumbnailByID(
			[In, MarshalAs(UnmanagedType.Struct)] WTS_THUMBNAILID thumbnailID,
			[In] uint cxyRequestedThumbSize,
			[Out][MarshalAs(UnmanagedType.Interface)] out ISharedBitmap ppvThumb,
			[Out] out WTS_CACHEFLAGS pOutFlags
		);
	}
}
