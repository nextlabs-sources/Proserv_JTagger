package com.nextlabs.jtagger;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.hpsf.CustomProperties;
import org.apache.poi.hpsf.CustomProperty;
import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.MarkUnsupportedException;
import org.apache.poi.hpsf.NoPropertySetStreamException;
import org.apache.poi.hpsf.PropertySet;
import org.apache.poi.hpsf.PropertySetFactory;
import org.apache.poi.hpsf.UnexpectedPropertySetTypeException;
import org.apache.poi.hpsf.WritingNotSupportedException;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class Office2k3Tagger extends AbstractTagger {

	private POIFSFileSystem poifs;
	DirectoryEntry directoryEntry;
	DocumentSummaryInformation documentSummaryInformation;
	CustomProperties customProperties;

	public Office2k3Tagger(String filename) throws IOException {
		super(filename);

		InputStream is = null;

		try {
			is = new FileInputStream(filename);
			poifs = new POIFSFileSystem(is);
		} finally {
			if (is != null) {
				is.close();
			}
		}

		try {
			directoryEntry = poifs.getRoot();
			documentSummaryInformation = getDocumentSummaryInformation(directoryEntry);
			customProperties = documentSummaryInformation.getCustomProperties() == null ? new CustomProperties() : documentSummaryInformation.getCustomProperties();
		} catch (Exception err) {
			throw new IOException(err);
		}
	}

	public Office2k3Tagger(String filename, InputStream is) throws IOException {
		try {
			poifs = new POIFSFileSystem(is);
		} finally {
			if (is != null) {
				is.close();
			}
		}

		try {
			directoryEntry = poifs.getRoot();
			documentSummaryInformation = getDocumentSummaryInformation(directoryEntry);
			customProperties = documentSummaryInformation.getCustomProperties() == null ? new CustomProperties() : documentSummaryInformation.getCustomProperties();
		} catch (Exception err) {
			throw new IOException(err);
		}
	}

	@Override
	public HashMap<String, Object> getAllTags() throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException {
		return (getAllTags(true));
	}

	@Override
	public HashMap<String, Object> getAllTags(boolean bToLowerCase) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException {
		HashMap<String, Object> tags = new HashMap<String, Object>();

		for (Entry<Object, CustomProperty> entry : customProperties.entrySet()) {
			if (bToLowerCase) tags.put(entry.getValue().getName().toLowerCase(), entry.getValue().getValue());
			else tags.put(entry.getValue().getName(), entry.getValue().getValue());
		}

		return tags;
	}

	@Override
	public boolean deleteAllTags() throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		return deleteAllTags(false);
	}

	@Override
	public boolean deleteAllTags(boolean editSecured) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		boolean flag = false;

		customProperties.clear();
		documentSummaryInformation.setCustomProperties(customProperties);
		documentSummaryInformation.write(directoryEntry, DocumentSummaryInformation.DEFAULT_STREAM_NAME);
		flag = true;

		return flag;
	}

	@Override
	public Object getTag(String key) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException {
		for (Entry<Object, CustomProperty> entry : customProperties.entrySet()) {
			if (key.equalsIgnoreCase(entry.getValue().getName())) {
				return entry.getValue().getValue();
			}
		}

		return null;
	}

	@Override
	public boolean addTag(String key, Object value) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		return addTag(key, value, false);
	}

	@Override
	public boolean addTag(String key, Object value, boolean editSecure) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		boolean flag = true;

		for (Entry<Object, CustomProperty> property : customProperties.entrySet()) {
			if (property.getValue().getName().equalsIgnoreCase(key) && property.getValue().getValue().equals(value)) {
				flag = false;
				break;
			}
		}

		if (flag) {
			customProperties.put(key, value.toString());
			documentSummaryInformation.setCustomProperties(customProperties);
			documentSummaryInformation.write(directoryEntry, DocumentSummaryInformation.DEFAULT_STREAM_NAME);
		}

		return flag;
	}

	@Override
	public boolean addTags(Map<String, Object> tags) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		return addTags(tags, false);
	}

	@Override
	public boolean addTags(Map<String, Object> tags, boolean editSecured) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		boolean flag = false;

		if (tags != null) {
			for (Entry<String, Object> tag : tags.entrySet()) {
				boolean needToAdd = true;

				for (Entry<Object, CustomProperty> property : customProperties.entrySet()) {
					if (property.getValue().getName().equalsIgnoreCase(tag.getKey()) && property.getValue().getValue().equals(tag.getValue())) {
						needToAdd = false;
						break;
					}
				}

				if (needToAdd) {
					customProperties.put(tag.getKey(), tag.getValue().toString());
					flag = true;
				}
			}

			if (flag) {
				documentSummaryInformation.setCustomProperties(customProperties);
				documentSummaryInformation.write(directoryEntry, DocumentSummaryInformation.DEFAULT_STREAM_NAME);
			}
		}

		return flag;
	}

	@Override
	public boolean deleteTag(String key) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		return deleteTag(key, false);
	}

	@Override
	public boolean deleteTag(String key, boolean editSecured) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		boolean flag = false;

		for (Object cProperties : customProperties.keySet()) {
			if (((String) cProperties).equalsIgnoreCase(key)) {
				customProperties.remove((String) cProperties);
				documentSummaryInformation.setCustomProperties(customProperties);
				documentSummaryInformation.write(directoryEntry, DocumentSummaryInformation.DEFAULT_STREAM_NAME);
				flag = true;

				break;
			}
		}

		return flag;
	}

	@Override
	public boolean deleteTags(Collection<String> tags) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		return deleteTags(tags, false);
	}

	@Override
	public boolean deleteTags(Collection<String> tags, boolean editSecured) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		boolean flag = false;

		Iterator<Map.Entry<Object, CustomProperty>> iterator = customProperties.entrySet().iterator();

		while (iterator.hasNext()) {
			Entry<Object, CustomProperty> property = iterator.next();

			for (String key : tags) {
				if (property.getValue().getName().equalsIgnoreCase(key)) {
					iterator.remove();
					flag = true;

					break;
				}
			}
		}

		if (flag) {
			documentSummaryInformation.setCustomProperties(customProperties);
			documentSummaryInformation.write(directoryEntry, DocumentSummaryInformation.DEFAULT_STREAM_NAME);
		}

		return flag;
	}

	@Override
	public void save(String filename) throws FileNotFoundException, IOException {
		OutputStream outputStream = null;

		try {
			outputStream = new FileOutputStream(new File(filename));
			poifs.writeFilesystem(outputStream);
		} finally {
			if (outputStream != null) {
				outputStream.close();
			}
		}
	}
	
	@Override
	public void save(OutputStream os) throws FileNotFoundException, IOException {
		try {
			poifs.writeFilesystem(os);
		} finally {
			if (os != null) {
				os.close();
			}
		}
	}

	@Override
	public void close() throws IOException {
		// Nothing to close
	}

	private DocumentSummaryInformation getDocumentSummaryInformation(DirectoryEntry directoryEntry) throws IOException, MarkUnsupportedException, UnexpectedPropertySetTypeException, NoPropertySetStreamException {
		try {
			DocumentEntry documentEntry = (DocumentEntry) directoryEntry.getEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME);
			DocumentInputStream documentInputStream = new DocumentInputStream(documentEntry);
			PropertySet propertySet = new PropertySet(documentInputStream);
			documentInputStream.close();
			return new DocumentSummaryInformation(propertySet);
		} catch (FileNotFoundException e) {
			return PropertySetFactory.newDocumentSummaryInformation();
		}
	}
}
