package com.nextlabs.jtagger;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.xmlbeans.XmlException;

public class TaggerFactory {

	public static Tagger getTagger(String filename) throws IOException, OpenXML4JException, XmlException {
		
		String fileExtension = "";
		
		int i = filename.lastIndexOf('.');
		if (i > 0) {
			fileExtension = filename.substring(i+1).toLowerCase();
		}
		
		if(!fileExtension.equals("")) {
			if(fileExtension.equals("doc") || fileExtension.equals("ppt") || fileExtension.equals("xls")|| fileExtension.equals("msg") || fileExtension.equals("vsd")) {
				return new Office2k3Tagger(filename);
			} else if(fileExtension.equals("docx") || fileExtension.equals("pptx") || fileExtension.equals("xlsx")) {
				return new Office2k7Tagger(filename);
			} else if(fileExtension.equals("pdf")) {
				return new PDFTagger(filename);
			} else {
				return null;
			}
		}
		
		return null;
		
	}
	
	public static Tagger getTagger(String filename, InputStream is) throws IOException, OpenXML4JException, XmlException {
			
			String fileExtension = "";
			
			int i = filename.lastIndexOf('.');
			if (i > 0) {
				fileExtension = filename.substring(i+1).toLowerCase();
			}
			
			if(!fileExtension.equals("")) {
				if(fileExtension.equals("doc") || fileExtension.equals("ppt") || fileExtension.equals("xls")|| fileExtension.equals("msg") || fileExtension.equals("vsd")) {
					return new Office2k3Tagger(filename, is);
				} else if(fileExtension.equals("docx") || fileExtension.equals("pptx") || fileExtension.equals("xlsx")) {
					return new Office2k7Tagger(filename, is);
				} else if(fileExtension.equals("pdf")) {
					return new PDFTagger(filename, is);
				} else {
					return null;
				}
			}
			
			return null;
			
		}
	
}
