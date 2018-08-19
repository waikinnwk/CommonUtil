package com.common.util;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.GZIPInputStream;
import java.util.zip.GZIPOutputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

public class ZipUtil {
	public static byte[] compress(byte[] ba) throws Exception {
		ByteArrayOutputStream bao = new ByteArrayOutputStream();
		GZIPOutputStream gos = new GZIPOutputStream(bao);

		try {
			gos.write(ba);
		} finally {
			gos.close();
		}

		return bao.toByteArray();
	}

	public static byte[] decompress(byte[] bytes) throws Exception {
		StringBuffer sb = new StringBuffer();
		GZIPInputStream gis = new GZIPInputStream(new ByteArrayInputStream(bytes));

		try {
			byte[] buff = new byte[1024];
			int count = 0;
			while ((count = gis.read(buff, 0, buff.length)) > 0) {
				sb.append(new String(buff, 0, count));
			}
		} finally {
			gis.close();
		}

		return sb.toString().getBytes();
	}

	public static void compress(List<File> fileList, File outputFile) throws Exception {
		ZipOutputStream out = null;
		try {
			BufferedInputStream origin = null;
			FileOutputStream dest = new FileOutputStream(outputFile);
			out = new ZipOutputStream(new BufferedOutputStream(dest));
			// out.setMethod(ZipOutputStream.DEFLATED);
			byte data[] = new byte[1024];
			// get a list of files from current directory

			for (int i = 0; i < fileList.size(); i++) {
				FileInputStream fi = new FileInputStream(fileList.get(i));
				origin = new BufferedInputStream(fi, 1024);
				ZipEntry entry = new ZipEntry(fileList.get(i).getName());
				entry.setSize(fileList.get(i).length());
				out.putNextEntry(entry);
				int count;
				while ((count = origin.read(data, 0, 1024)) != -1) {
					out.write(data, 0, count);
				}
				origin.close();
			}
		} finally {
			out.finish();
			out.close();
		}
	}

	public static List<String> getZipEntryNameList(File zippedFile) throws Exception {
		FileInputStream fis = new FileInputStream(zippedFile);
		List<String> resultList = getZipEntryNameList(fis);
		fis.close();
		return resultList;
	}

	public static List<String> getZipEntryNameList(InputStream is) throws Exception {
		ZipInputStream zis = null;
		List<String> resultList = new ArrayList<String>();
		try {
			zis = new ZipInputStream(new BufferedInputStream(is));
			ZipEntry entry;
			while ((entry = zis.getNextEntry()) != null) {
				resultList.add(entry.getName());
			}
		} finally {
			zis.close();
		}
		return resultList;
	}

	public static List<File> decompressToFile(File zippedFile) throws Exception {
		FileInputStream fis = new FileInputStream(zippedFile);
		List<File> resultList = decompressToFile(fis);
		fis.close();
		return resultList;
	}

	public static List<File> decompressToFile(InputStream is) throws Exception {
		ZipInputStream zis = null;
		List<File> resultList = new ArrayList<File>();
		try {
			BufferedOutputStream dest = null;
			zis = new ZipInputStream(new BufferedInputStream(is));
			ZipEntry entry;
			while ((entry = zis.getNextEntry()) != null) {
				int count;
				byte data[] = new byte[1024];
				// write the files to the disk
				FileOutputStream fos = new FileOutputStream(entry.getName());
				dest = new BufferedOutputStream(fos, 1024);
				while ((count = zis.read(data, 0, 1024)) != -1) {
					dest.write(data, 0, count);
				}
				dest.flush();
				dest.close();
				resultList.add(new File(entry.getName()));
			}
		} finally {
			zis.close();
		}
		return resultList;
	}

	public static List<byte[]> decompressToByteArr(File zippedFile) throws Exception {
		FileInputStream fis = new FileInputStream(zippedFile);
		List<byte[]> resultList = decompressToByteArr(fis);
		fis.close();
		return resultList;
	}

	public static List<byte[]> decompressToByteArr(InputStream is) throws Exception {
		ZipInputStream zis = null;
		List<byte[]> resultList = new ArrayList<byte[]>();
		try {
			zis = new ZipInputStream(new BufferedInputStream(is));
			ZipEntry entry;
			while ((entry = zis.getNextEntry()) != null) {
				int count;
				byte buffer[] = new byte[1024];
				ByteArrayOutputStream baos = new ByteArrayOutputStream();
				// byte data[] = new byte[(int) entry.getSize()];
				// write the files to the disk
				// FileOutputStream fos = new FileOutputStream(entry.getName());
				// dest = new BufferedOutputStream(fos, 1024);
				// int index=0;
				while ((count = zis.read(buffer)) != -1) {

					// data[index+i] = buffer[i];
					baos.write(buffer,0,count);

					// index += count;
				}		
				resultList.add(baos.toByteArray());
				baos.close();
			}
		} finally {
			zis.close();
		}
		return resultList;
	}

}
