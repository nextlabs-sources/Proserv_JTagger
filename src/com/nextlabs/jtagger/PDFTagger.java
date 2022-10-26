package com.nextlabs.jtagger;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.TreeSet;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentInformation;
import org.apache.poi.hpsf.MarkUnsupportedException;
import org.apache.poi.hpsf.NoPropertySetStreamException;
import org.apache.poi.hpsf.UnexpectedPropertySetTypeException;
import org.apache.poi.hpsf.WritingNotSupportedException;

public class PDFTagger extends AbstractTagger {

	private PDDocument document;
	private static final Set<String> restrictedKeys;

	static {
		restrictedKeys = new TreeSet<String>(String.CASE_INSENSITIVE_ORDER);
		restrictedKeys.add("Title");
		restrictedKeys.add("Author");
		restrictedKeys.add("Subject");
		restrictedKeys.add("Keywords");
		restrictedKeys.add("Creator");
		restrictedKeys.add("Producer");
		restrictedKeys.add("CreationDate");
		restrictedKeys.add("ModDate");
		restrictedKeys.add("Trapped");
	}

	public PDFTagger(String filename) throws IOException {
		super(filename);

		document = PDDocument.load(new File(filename));
	}
	
	public PDFTagger(String filename, InputStream is) throws IOException {
		super(filename);

		document = PDDocument.load(is);
	}

	@Override
	public HashMap<String, Object> getAllTags() throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException {
		return (getAllTags(true));
	}

	@Override
	public HashMap<String, Object> getAllTags(boolean bToLowerCase) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException {
		HashMap<String, Object> tags = new HashMap<String, Object>();

		for (String key : document.getDocumentInformation().getMetadataKeys()) {
			if (!restrictedKeys.contains(key)) {
				if (bToLowerCase) tags.put(key.toLowerCase(), document.getDocumentInformation().getCustomMetadataValue(key));
				else tags.put(key, document.getDocumentInformation().getCustomMetadataValue(key));
			}
		}

		return tags;
	}

	@Override
	public boolean deleteAllTags() throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		return deleteAllTags(false);
	}

	@Override
	public boolean deleteAllTags(boolean editSecured) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		if (document.isEncrypted()) {
			if (!editSecured) return false;

			document.setAllSecurityToBeRemoved(true);
		}

		PDDocumentInformation newDocumentInformation = new PDDocumentInformation();
		boolean flag = false;

		for (String s : document.getDocumentInformation().getMetadataKeys()) {
			if (s.equalsIgnoreCase("Title")) {
				newDocumentInformation.setTitle(document.getDocumentInformation().getTitle());
			} else if (s.equalsIgnoreCase("Author")) {
				newDocumentInformation.setAuthor(document.getDocumentInformation().getAuthor());
			} else if (s.equalsIgnoreCase("Subject")) {
				newDocumentInformation.setSubject(document.getDocumentInformation().getSubject());
			} else if (s.equalsIgnoreCase("Keywords")) {
				newDocumentInformation.setKeywords(document.getDocumentInformation().getKeywords());
			} else if (s.equalsIgnoreCase("Creator")) {
				newDocumentInformation.setCreator(document.getDocumentInformation().getCreator());
			} else if (s.equalsIgnoreCase("Producer")) {
				newDocumentInformation.setProducer(document.getDocumentInformation().getProducer());
			} else if (s.equalsIgnoreCase("CreationDate")) {
				newDocumentInformation.setCreationDate(document.getDocumentInformation().getCreationDate());
			} else if (s.equalsIgnoreCase("ModDate")) {
				newDocumentInformation.setModificationDate(document.getDocumentInformation().getModificationDate());
			} else if (s.equalsIgnoreCase("Trapped")) {
				newDocumentInformation.setTrapped(document.getDocumentInformation().getTrapped());
			}
		}

		document.setDocumentInformation(newDocumentInformation);
		flag = true;

		return flag;
	}

	@Override
	public Object getTag(String key) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException {
		if (!restrictedKeys.contains(key)) {
			// For case insensitive key handling
			Iterator<String> iterator = document.getDocumentInformation().getMetadataKeys().iterator();
			while (iterator.hasNext()) {
				String sKey = iterator.next();
				if (sKey.equalsIgnoreCase(key)) {
					return (document.getDocumentInformation().getCustomMetadataValue(sKey));
				}
			}
		}

		return null;
	}

	@Override
	public boolean addTag(String key, Object value) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		return addTag(key, value, false);
	}

	@Override
	public boolean addTag(String key, Object value, boolean editSecured) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		if (document.isEncrypted()) {
			if (!editSecured) return false;

			document.setAllSecurityToBeRemoved(true);
		}

		boolean flag = true;

		if (restrictedKeys.contains(key)) {
			throw new IllegalArgumentException("Key cannot be a restricted parameter.");
		}

		for (String metadataKey : document.getDocumentInformation().getMetadataKeys()) {
			if (metadataKey.equalsIgnoreCase(key) && document.getDocumentInformation().getCustomMetadataValue(metadataKey).equals(value)) {
				flag = false;
				break;
			}
		}

		if (flag) {
			document.getDocumentInformation().setCustomMetadataValue(key, value.toString());
		}

		return flag;
	}

	@Override
	public boolean addTags(Map<String, Object> tags) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		return addTags(tags, false);
	}

	@Override
	public boolean addTags(Map<String, Object> tags, boolean editSecured) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		if (document.isEncrypted()) {
			if (!editSecured) return false;

			document.setAllSecurityToBeRemoved(true);
		}

		boolean flag = false;

		if (tags != null) {
			for (Entry<String, Object> tag : tags.entrySet()) {
				if (restrictedKeys.contains(tag.getKey())) {
					throw new IllegalArgumentException("Key cannot be a restricted parameter.");
				}

				boolean needToAdd = true;
				for (String metadataKey : document.getDocumentInformation().getMetadataKeys()) {
					if (metadataKey.equalsIgnoreCase(tag.getKey()) && document.getDocumentInformation().getCustomMetadataValue(metadataKey).equals(tag.getValue())) {
						needToAdd = false;
						break;
					}
				}

				if (needToAdd) {
					document.getDocumentInformation().setCustomMetadataValue(tag.getKey(), tag.getValue().toString());
					flag = true;
				}
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
		if (restrictedKeys.contains(key)) {
			throw new IllegalArgumentException("Cannot remove restricted parameter.");
		}

		if (document.isEncrypted()) {
			if (!editSecured) return false;

			document.setAllSecurityToBeRemoved(true);
		}

		PDDocumentInformation newDocumentInformation = new PDDocumentInformation();
		boolean flag = false;

		for (String s : document.getDocumentInformation().getMetadataKeys()) {
			if (s.equalsIgnoreCase("Title")) {
				newDocumentInformation.setTitle(document.getDocumentInformation().getTitle());
			} else if (s.equalsIgnoreCase("Author")) {
				newDocumentInformation.setAuthor(document.getDocumentInformation().getAuthor());
			} else if (s.equalsIgnoreCase("Subject")) {
				newDocumentInformation.setSubject(document.getDocumentInformation().getSubject());
			} else if (s.equalsIgnoreCase("Keywords")) {
				newDocumentInformation.setKeywords(document.getDocumentInformation().getKeywords());
			} else if (s.equalsIgnoreCase("Creator")) {
				newDocumentInformation.setCreator(document.getDocumentInformation().getCreator());
			} else if (s.equalsIgnoreCase("Producer")) {
				newDocumentInformation.setProducer(document.getDocumentInformation().getProducer());
			} else if (s.equalsIgnoreCase("CreationDate")) {
				newDocumentInformation.setCreationDate(document.getDocumentInformation().getCreationDate());
			} else if (s.equalsIgnoreCase("ModDate")) {
				newDocumentInformation.setModificationDate(document.getDocumentInformation().getModificationDate());
			} else if (s.equalsIgnoreCase("Trapped")) {
				newDocumentInformation.setTrapped(document.getDocumentInformation().getTrapped());
			} else {
				if (!s.equalsIgnoreCase(key)) {
					newDocumentInformation.setCustomMetadataValue(s, document.getDocumentInformation().getCustomMetadataValue(s));
				} else {
					flag = true;
				}
			}
		}

		document.setDocumentInformation(newDocumentInformation);

		return flag;
	}

	@Override
	public boolean deleteTags(Collection<String> keys) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		return deleteTags(keys, false);
	}

	@Override
	public boolean deleteTags(Collection<String> keys, boolean editSecured) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		if (document.isEncrypted()) {
			if (!editSecured) return false;

			document.setAllSecurityToBeRemoved(true);
		}

		PDDocumentInformation newDocumentInformation = new PDDocumentInformation();
		boolean flag = false;

		for (String s : document.getDocumentInformation().getMetadataKeys()) {
			if (s.equalsIgnoreCase("Title")) {
				newDocumentInformation.setTitle(document.getDocumentInformation().getTitle());
			} else if (s.equalsIgnoreCase("Author")) {
				newDocumentInformation.setAuthor(document.getDocumentInformation().getAuthor());
			} else if (s.equalsIgnoreCase("Subject")) {
				newDocumentInformation.setSubject(document.getDocumentInformation().getSubject());
			} else if (s.equalsIgnoreCase("Keywords")) {
				newDocumentInformation.setKeywords(document.getDocumentInformation().getKeywords());
			} else if (s.equalsIgnoreCase("Creator")) {
				newDocumentInformation.setCreator(document.getDocumentInformation().getCreator());
			} else if (s.equalsIgnoreCase("Producer")) {
				newDocumentInformation.setProducer(document.getDocumentInformation().getProducer());
			} else if (s.equalsIgnoreCase("CreationDate")) {
				newDocumentInformation.setCreationDate(document.getDocumentInformation().getCreationDate());
			} else if (s.equalsIgnoreCase("ModDate")) {
				newDocumentInformation.setModificationDate(document.getDocumentInformation().getModificationDate());
			} else if (s.equalsIgnoreCase("Trapped")) {
				newDocumentInformation.setTrapped(document.getDocumentInformation().getTrapped());
			} else {
				Iterator<String> iterator = keys.iterator();
				boolean removeKey = false;
				while (iterator.hasNext()) {
					if (s.equalsIgnoreCase(iterator.next())) {
						removeKey = true;
						flag = true;
						break;
					}
				}

				if (!removeKey) {
					newDocumentInformation.setCustomMetadataValue(s, document.getDocumentInformation().getCustomMetadataValue(s));
				}
			}
		}

		document.setDocumentInformation(newDocumentInformation);

		return flag;
	}

	@Override
	public void save(String filename) throws FileNotFoundException, IOException {
		document.save(filename);
		document.close();
	}
	
	@Override
	public void save(OutputStream os) throws FileNotFoundException, IOException {
		document.save(os);
		document.close();
	}

	@Override
	public void close() throws IOException {
		if (document != null) {
			document.close();
		}
	}
}
