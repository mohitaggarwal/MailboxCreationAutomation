using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuickXORHashTest
{
	class Program
	{
		static void Main(string[] args)
		{
			byte[] bytes1 = Encoding.ASCII.GetBytes("abc");
			byte[] bytes2 = Encoding.ASCII.GetBytes("def");
			byte[] bytes3 = Encoding.ASCII.GetBytes("abcdeff");
			QuickXorHash quickXorHash1 = new QuickXorHash();
			quickXorHash1.HashCore(bytes1, 0, bytes1.Length);
			quickXorHash1.HashCore(bytes2, 0, bytes2.Length);
			string hashInChunks = Convert.ToBase64String(quickXorHash1.HashFinal());

			QuickXorHash quickXorHash2 = new QuickXorHash();
			quickXorHash2.HashCore(bytes3, 0, bytes3.Length);
			string hashFull = Convert.ToBase64String(quickXorHash2.HashFinal());

			Console.WriteLine($"HashInChunks: {hashInChunks}, HashFull: {hashFull}, IsEqual: {hashFull.Equals(hashInChunks)}");
		}
	}
}
