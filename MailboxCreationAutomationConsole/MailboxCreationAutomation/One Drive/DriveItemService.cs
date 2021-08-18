using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace MailboxCreationAutomation
{
	public class DriveItemService
	{
		private GraphServiceWrapper _GraphServiceWrapper;
		private string _userId;

		public DriveItemService(GraphServiceWrapper graphServiceWrapper, string userId)
		{
			_GraphServiceWrapper = graphServiceWrapper;
			_userId = userId;
		}

		private string GetInnerExceptionMessage(Exception ex, string message)
		{
			if (ex == null)
				return message;
			else
				return GetInnerExceptionMessage(ex.InnerException, $"{message}\n{ex.Message}");
		}

		private string GetToken(string link, string tokenName)
		{
			if (string.IsNullOrEmpty(link))
				return "";
			Uri uri = new Uri(link);
			var queries = HttpUtility.ParseQueryString(uri.Query);
			return queries.Get(tokenName);
		}

		private void CreateFile(string fileName, long sizeInKB)
		{
			using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None))
			{
				fs.SetLength(sizeInKB * 1024);
			}
		}

		private UploadSession GetUploadSession(string filePath)
		{
			// Use properties to specify the conflict behavior
			// in this case, replace
			var uploadProps = new DriveItemUploadableProperties
			{
				ODataType = null,
				AdditionalData = new Dictionary<string, object>
					{
						{ "@microsoft.graph.conflictBehavior", "replace" }
					}
			};
			// Create the upload session
			// itemPath does not need to be a path to an existing item
			return _GraphServiceWrapper.GraphServiceClient
					.Users[_userId].Drive.Root
					.ItemWithPath(filePath)
					.CreateUploadSession(uploadProps)
					.Request()
					.PostAsync()
					.Result;
		}

		public List<DriveItemInfo> GetDriveItems( string driveId, ref string skipToken, string pageNumber = "100")
		{
			var queryOptions = new List<QueryOption>()
			{
				new QueryOption("$top", pageNumber)
			};
			//string skipToken = GetToken(nextLink, "token");
			if (!string.IsNullOrEmpty(skipToken))
			{
				queryOptions.Add(new QueryOption("$skiptoken", skipToken));
			}

			//string deltaToken = GetToken(deltaLink, "token");
			//if (!string.IsNullOrEmpty(deltaToken))
			//{
			//	queryOptions.Add(new QueryOption("$deltatoken", deltaToken));
			//}

			var driveItems = _GraphServiceWrapper.GraphServiceClient
								.Users[_userId]
								.Drive
								.Root
								.Delta()
								.Request(queryOptions)
								.GetAsync().Result;

			skipToken = driveItems
							.NextPageRequest?
							.QueryOptions?
							.FirstOrDefault(
								x => string.Equals("$skiptoken", x.Name, StringComparison.InvariantCultureIgnoreCase))?
							.Value;
			//Object deltaLinkObj;
			//deltaLink = string.Empty;
			//if (driveItems.AdditionalData.TryGetValue("@odata.deltaLink", out deltaLinkObj))
			//{
			//	deltaLink = deltaLinkObj as string;
			//}

			//Object nextLinkObj;
			//nextLink = string.Empty;
			//if (driveItems.AdditionalData.TryGetValue("@odata.nextLink", out nextLinkObj))
			//{
			//	nextLink = nextLinkObj as string;
			//}
			return driveItems.Select(
				x => new DriveItemInfo {
					Id = x.Id,
					Name = x.Name,
					Size = x.Size,
					DriveItemType = x.Folder != null ? DriveItemType.Folder
									: (x.File != null ? DriveItemType.File : DriveItemType.Other)
				})
				.ToList<DriveItemInfo>();
		}

		public DriveItemInfo GetFileItem(string itemId)
		{
			string downloadUrl = string.Empty;
			var driveItem = _GraphServiceWrapper.GraphServiceClient
								.Users[_userId]
								.Drive
								.Items[itemId]
								.Request()
								.GetAsync().Result;
			if (driveItem != null)
			{
				object downloadUrlObj;
				driveItem.AdditionalData.TryGetValue("@microsoft.graph.downloadUrl", out downloadUrlObj);
				downloadUrl = downloadUrlObj as string;
			}
			return new DriveItemInfo
			{
				Id = driveItem.Id,
				Name = driveItem.Name,
				Size = driveItem.Size,
				DriveItemType = driveItem.Folder != null ? DriveItemType.Folder
									: (driveItem.File != null ? DriveItemType.File : DriveItemType.Other),
				CheckSum = driveItem.File?.Hashes.QuickXorHash,
				DownloadUrl = downloadUrl
			};
		}

		public byte[] GetFileChunk(string downloadUrl, long offset, long chunkSize)
		{
			byte[] bytesInStream = null;
			bool needRetry = true;
			int retryCount = 0;
			do
			{
				try
				{
					// Create the request message with the download URL and Range header.
					using (HttpRequestMessage req = new HttpRequestMessage(HttpMethod.Get, (string)downloadUrl))
					{
						req.Headers.Range = new RangeHeaderValue(offset, chunkSize + offset - 1);

						// We can use the the client library to send this although it does add an authentication cost.
						// HttpResponseMessage response = await graphClient.HttpProvider.SendAsync(req);
						// Since the download URL is preauthenticated, and we aren't deserializing objects, 
						// we'd be better to make the request with HttpClient.
						using (var client = new HttpClient())
						{
							using (HttpResponseMessage response = client.SendAsync(req).Result)
							{

								using (Stream responseStream = response.Content.ReadAsStreamAsync().Result)
								{
									bytesInStream = new byte[chunkSize];
									int read = responseStream.Read(bytesInStream, 0, (int)bytesInStream.Length);
									Console.WriteLine($"Number of bytes read: {read}");
								}
							}
						}
					}
					needRetry = false;
				}
				catch (Exception ex)
				{
					string message = GetInnerExceptionMessage(ex, "");
					Console.WriteLine($"Exception occur while downloading file chunk. offset: {offset}, chunkSize: {chunkSize}. Detail: {message}");
					Logger.FileLogger.Error($"Exception occur while downloading file chunk. offset: {offset}, chunkSize: {chunkSize}. Detail: {message}");
					//Console.WriteLine($"Retrying after {EWSServiceConstants.RETRY_AFTER / 1000}sec");
					//Logger.FileLogger.Warning($"Retrying after {EWSServiceConstants.RETRY_AFTER / 1000}sec");
					//Thread.Sleep(EWSServiceConstants.RETRY_AFTER);
					needRetry = true;
					retryCount++;
				}
			} while (needRetry && retryCount < EWSServiceConstants.RETRY_COUNT);

			return bytesInStream;
		}

		public async Task DownloadFile(string itemId, string filePath)
		{
			using (var contentStream = await _GraphServiceWrapper.GraphServiceClient.Users[_userId].Drive.Items[itemId].Content.Request().GetAsync())
			{
				byte[] byteBuffer = new byte[4096];
				using (System.IO.FileStream output = new FileStream(filePath, FileMode.Create))
				{
					await contentStream.CopyToAsync(output);
				}
			}
		}

		public string DownloadFileByChunks(string itemId, string filePath, long chunkSize)
		{
			QuickXorHash quickXorHash = new QuickXorHash();
			var driveItem = GetFileItem(itemId);
			if (driveItem != null)
			{
				if (!System.IO.File.Exists(filePath))
				{
					var fileStream = System.IO.File.Create(filePath);
					fileStream.Close();
				}
				string downloadUrl = driveItem.DownloadUrl;
				//const long DefaultChunkSize = 5 * 1024 * 1024; // 50 KB, TODO: change chunk size to make it realistic for a large file.
				long ChunkSize = chunkSize;
				long offset = 0;
				long size = (long)driveItem.Size;
				int numberOfChunks = Convert.ToInt32(size / chunkSize);
				// We are incrementing the offset cursor after writing the response stream to a file after each chunk. 
				// Subtracting one since the size is 1 based, and the range is 0 base. There should be a better way to do
				// this but I haven't spent the time on that.
				int lastChunkSize = Convert.ToInt32(size % chunkSize);
				if (lastChunkSize > 0) { numberOfChunks++; }
				Console.WriteLine($"Chunk size in bytes: {chunkSize}.");
				Logger.FileLogger.Info($"Chunk size in bytes: {chunkSize}.");
				for (int i = 0; i < numberOfChunks; i++)
				{
					// Setup the last chunk to request. This will be called at the end of this loop.
					if (i == numberOfChunks - 1 && lastChunkSize > 0)
					{
						ChunkSize = lastChunkSize;
					}
					Console.WriteLine($"OffSet: {offset}, Chunk size: {ChunkSize}");
					var bytesDownloaded = GetFileChunk(downloadUrl, offset, ChunkSize);
					if (bytesDownloaded?.Length > 0)
					{
						using (FileStream fileStream = new FileStream(filePath, FileMode.Append))
						{
							fileStream.Write(bytesDownloaded, 0, bytesDownloaded.Length);
						}
						quickXorHash.TransformBlock(bytesDownloaded, 0, bytesDownloaded.Length, bytesDownloaded, 0);
					}
					offset += ChunkSize; // Move the offset cursor to the next chunk.
				}
				quickXorHash.TransformFinalBlock(new byte[0], 0, 0);
				return Convert.ToBase64String(quickXorHash.Hash);
			}
			else
			{
				Console.WriteLine($"Item: {itemId} not found.");
				Logger.FileLogger.Warning($"Item: {itemId} not found.");
				return "";
			}
		}

		public void DownloadFileByChunks(string itemId, string filePath, long chunkSizeInBytes, int parallelThreads=5)
		{
			var driveItem = GetFileItem(itemId);
			if (driveItem != null)
			{
				if (!System.IO.File.Exists(filePath))
				{
					var fileStream = System.IO.File.Create(filePath);
					fileStream.Close();
				}
				string downloadUrl = driveItem.DownloadUrl;
				long chunkSize = chunkSizeInBytes;
				long offset = 0;
				long size = (long)driveItem.Size;
				int numberOfChunks = Convert.ToInt32(size / chunkSize);
				// We are incrementing the offset cursor after writing the response stream to a file after each chunk. 
				// Subtracting one since the size is 1 based, and the range is 0 base. There should be a better way to do
				// this but I haven't spent the time on that.
				int lastChunkSize = Convert.ToInt32(size % chunkSize);

				int numberOfThreadsCall = numberOfChunks / parallelThreads;
				int lastThreadCall = numberOfChunks % parallelThreads;
				if (lastThreadCall > 0)
					numberOfThreadsCall++;
				for(int i=0;i< numberOfThreadsCall;i++)
				{
					if(i == numberOfThreadsCall -1 && lastThreadCall > 0)
					{
						parallelThreads = lastThreadCall;
					}
					SortedDictionary<int, byte[]> data = new SortedDictionary<int, byte[]>();
					Parallel.For(0, parallelThreads, (j) =>
					  {
						  data[j] = GetFileChunk(downloadUrl, offset + j * chunkSize, chunkSize);
						  
					  });
					foreach (var bytesDownloaded in data)
					{
						if (bytesDownloaded.Value?.Length > 0)
						{
							using (FileStream fileStream = new FileStream(filePath, FileMode.Append))
							{
								fileStream.Write(bytesDownloaded.Value, 0, bytesDownloaded.Value.Length);
							}
						}
					}
					offset += parallelThreads * chunkSize;
				}
				if (lastChunkSize > 0)
				{
					var bytesDownloaded = GetFileChunk(downloadUrl, offset, lastChunkSize);
					if (bytesDownloaded?.Length > 0)
					{
						using (FileStream fileStream = new FileStream(filePath, FileMode.Append))
						{
							fileStream.Write(bytesDownloaded, 0, bytesDownloaded.Length);
						}
					}
				}
			}
		}

		public string UploadFile(string filePath, string odFilePath)
		{
			string itemId = string.Empty;
			using (var fileStream = System.IO.File.OpenRead(filePath))
			{
				var uploadSession = GetUploadSession(odFilePath);

				// Max slice size must be a multiple of 320 KiB
				int maxSliceSize = 16 * 320 * 1024;
				var fileUploadTask =
					new LargeFileUploadTask<DriveItemInfo>(uploadSession, fileStream, maxSliceSize);

				// Create a callback that is invoked after each slice is uploaded
				IProgress<long> progress = new Progress<long>(prog => {
					Console.WriteLine($"Uploaded {prog} bytes of {fileStream.Length} bytes");
				});

				try
				{
					// Upload the file
					var uploadResult = fileUploadTask.UploadAsync(progress).Result;

					if (uploadResult.UploadSucceeded)
					{
						// The ItemResponse object in the result represents the
						// created item.
						Logger.FileLogger.Info($"Upload complete, item ID: {uploadResult.ItemResponse.Id}");
						itemId = uploadResult.ItemResponse.Id;
					}
					else
					{
						Logger.FileLogger.Error("Upload failed");
					}
				}
				catch (ServiceException ex)
				{
					Logger.FileLogger.Error($"Error uploading: {ex.ToString()}");
				}
			}
			return itemId;
		}

		private void CreatFiles(ODFileToCreate filesToCreate, string odFilePath, string prefix, int number)
		{
			string directory = $"{_userId}_{DateTime.Now.ToString("yyyy - MM - dd HH - mm - ss - fff")}";
			DirectoryInfo di = System.IO.Directory.CreateDirectory(directory);
			try
			{
				string fileName = $"{prefix}_{filesToCreate.Prefix}_File_{number}";
				string filePath = Path.Combine(directory, fileName);
				string odFilePathWithName = Path.Combine(odFilePath, fileName);
				CreateFile(filePath, filesToCreate.SizeInKB);
				string id = UploadFile(filePath, odFilePathWithName);
				Logger.FileLogger.Info($"File '{fileName}' and id '{id}' created successfully.");
			}
			finally
			{
				if (di.Exists)
				{
					di.Delete(true);
				}
			}
		}

		public void CreatFiles(ODFileToCreate filesToCreate, string odFilePath, string prefix)
		{
			if (filesToCreate != null)
			{
				for (int i = 1; i <= filesToCreate.Count; i++)
				{
					CreatFiles(filesToCreate, odFilePath, prefix, i);
				}
			}
		}
	}
}
