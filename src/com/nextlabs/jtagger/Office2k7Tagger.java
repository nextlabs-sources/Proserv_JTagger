package com.nextlabs.jtagger;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.POIXMLProperties.CustomProperties;
import org.apache.poi.hpsf.MarkUnsupportedException;
import org.apache.poi.hpsf.NoPropertySetStreamException;
import org.apache.poi.hpsf.UnexpectedPropertySetTypeException;
import org.apache.poi.hpsf.WritingNotSupportedException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xslf.XSLFSlideShow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.impl.piccolo.io.FileFormatException;
import org.openxmlformats.schemas.officeDocument.x2006.customProperties.CTProperty;

public class Office2k7Tagger extends AbstractTagger {

	private POIXMLDocument document;

	public Office2k7Tagger(String filename) throws IOException, OpenXML4JException, XmlException {
		super(filename);

		InputStream fs = null;

		try {
			fs = new FileInputStream(filename);

			if (filename.toLowerCase().endsWith("docx")) {
				document = new XWPFDocument(OPCPackage.open(fs));
			} else if (filename.toLowerCase().endsWith("xlsx")) {
				document = new XSSFWorkbook(OPCPackage.open(fs));
			} else if (filename.toLowerCase().endsWith("pptx")) {
				document = new XSLFSlideShow(OPCPackage.open(fs));
			} else {
				throw new FileFormatException("Unsupported file format for Tagger.");
			}
		} finally {
			if (fs != null) {
				fs.close();
			}
		}
	}
	
	public Office2k7Tagger(String filename, InputStream is) throws IOException, OpenXML4JException, XmlException {
		try {
			if (filename.toLowerCase().endsWith("docx")) {
				document = new XWPFDocument(OPCPackage.open(is));
			} else if (filename.toLowerCase().endsWith("xlsx")) {
				document = new XSSFWorkbook(OPCPackage.open(is));
			} else if (filename.toLowerCase().endsWith("pptx")) {
				document = new XSLFSlideShow(OPCPackage.open(is));
			} else {
				throw new FileFormatException("Unsupported file format for Tagger.");
			}
		} finally {
			if (is != null) {
				is.close();
			}
		}
	}

	@Override
	public HashMap<String, Object> getAllTags() {
		return (getAllTags(true));
	}

	@Override
	public boolean deleteAllTags() throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		return deleteAllTags(false);
	}

	@Override
	public boolean deleteAllTags(boolean editSecured) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		CustomProperties cp = document.getProperties().getCustomProperties();
		boolean flag = false;

		if (cp != null) {
			cp.getUnderlyingProperties().getPropertyList().clear();
			;
			flag = true;
		}

		return flag;
	}

	@Override
	public HashMap<String, Object> getAllTags(boolean bToLowerCase) {
		HashMap<String, Object> tags = new HashMap<String, Object>();
		CustomProperties cp = document.getProperties().getCustomProperties();
		if (cp != null) {
			for (CTProperty ctp : cp.getUnderlyingProperties().getPropertyList()) {
				// TODO: Support for string based tags only.
				if (bToLowerCase) tags.put(ctp.getName().toLowerCase(), getStringValue(ctp));
				else tags.put(ctp.getName(), getStringValue(ctp));
			}
		}

		return tags;
	}

	@Override
	public Object getTag(String key) {
		CustomProperties cp = document.getProperties().getCustomProperties();

		if (cp != null) {
			for (CTProperty ctp : cp.getUnderlyingProperties().getPropertyList()) {
				// TODO: Add support for non-string attributes.
				if (key.equalsIgnoreCase(ctp.getName())) {
					return getStringValue(ctp);
				}
			}
		}

		return null;
	}

	private String getStringValue(CTProperty property) {
		String value = "";

		if (property.isSetLpwstr()) {
			value = property.getLpwstr();
		} else if (property.isSetLpstr()) {
			value = property.getLpstr();
		} else if (property.isSetDate()) {
			value = property.getDate().toString();
		} else if (property.isSetFiletime()) {
			value = property.getFiletime().toString();
		} else if (property.isSetBool()) {
			value = Boolean.toString(property.getBool());
		}

		// Integers
		else if (property.isSetI1()) {
			value = Integer.toString(property.getI1());
		} else if (property.isSetI2()) {
			value = Integer.toString(property.getI2());
		} else if (property.isSetI4()) {
			value = Integer.toString(property.getI4());
		} else if (property.isSetI8()) {
			value = Long.toString(property.getI8());
		} else if (property.isSetInt()) {
			value = Integer.toString(property.getInt());
		}

		// Unsigned Integers
		else if (property.isSetUi1()) {
			value = Integer.toString(property.getUi1());
		} else if (property.isSetUi2()) {
			value = Integer.toString(property.getUi2());
		} else if (property.isSetUi4()) {
			value = Long.toString(property.getUi4());
		} else if (property.isSetUi8()) {
			value = property.getUi8().toString();
		} else if (property.isSetUint()) {
			value = Long.toString(property.getUint());
		}

		// Reals
		else if (property.isSetR4()) {
			value = Float.toString(property.getR4());
		} else if (property.isSetR8()) {
			value = Double.toString(property.getR8());
		} else if (property.isSetDecimal()) {
			BigDecimal d = property.getDecimal();
			if (d == null) {
				value = null;
			} else {
				value = d.toPlainString();
			}
		}

		else if (property.isSetArray()) {
			// TODO Fetch the array values and output
		} else if (property.isSetVector()) {
			// TODO Fetch the vector values and output
		}

		else if (property.isSetBlob() || property.isSetOblob()) {
			// TODO Decode, if possible
		} else if (property.isSetStream() || property.isSetOstream() || property.isSetVstream()) {
			// TODO Decode, if possible
		} else if (property.isSetStorage() || property.isSetOstorage()) {
			// TODO Decode, if possible
		}

		return value;
	}

	@Override
	public boolean addTag(String key, Object value) throws WritingNotSupportedException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, IOException {
		return addTag(key, value, false);
	}

	@Override
	public boolean addTag(String key, Object value, boolean editSecured) throws WritingNotSupportedException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, IOException {
		CustomProperties cp = document.getProperties().getCustomProperties();
		boolean flag = true;

		if (cp != null) {
			List<CTProperty> ctProperties = cp.getUnderlyingProperties().getPropertyList();
			@SuppressWarnings("deprecation")
			CTProperty[] ctPropertiesArray = cp.getUnderlyingProperties().getPropertyArray();

			for (int i = 0; i < ctPropertiesArray.length; i++) {
				if (ctPropertiesArray[i].getName().equalsIgnoreCase(key)) {
					if (getStringValue(ctPropertiesArray[i]).equals(value)) {
						flag = false;
					} else {
						// Delete the key if the same key exist in the current file
						ctProperties.remove(i);
					}
					break;
				}
			}

			if (flag) {
				cp.addProperty(key, value.toString());
			}
		}

		return flag;
	}

	@Override
	public boolean addTags(Map<String, Object> tags) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		return addTags(tags, false);
	}

	@Override
	public boolean addTags(Map<String, Object> tags, boolean editSecured) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		CustomProperties cp = document.getProperties().getCustomProperties();
		boolean flag = false;

		if (cp != null && tags != null) {
			for (Entry<String, Object> tag : tags.entrySet()) {
				// CTProperties need to be refresh after every loop
				List<CTProperty> ctProperties = cp.getUnderlyingProperties().getPropertyList();
				@SuppressWarnings("deprecation")
				CTProperty[] ctPropertiesArray = cp.getUnderlyingProperties().getPropertyArray();
				boolean addProperty = true;

				for (int i = 0; i < ctPropertiesArray.length; i++) {
					if (ctPropertiesArray[i].getName().equalsIgnoreCase(tag.getKey())) {
						if (getStringValue(ctPropertiesArray[i]).equals(tag.getValue())) {
							addProperty = false;
						} else {
							// Delete the key if the same key exist in the current file
							ctProperties.remove(i);
						}
						break;
					}
				}

				if (addProperty) {
					cp.addProperty(tag.getKey(), tag.getValue().toString());
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
		CustomProperties cp = document.getProperties().getCustomProperties();
		boolean flag = false;

		if (cp != null) {
			List<CTProperty> ctProperties = cp.getUnderlyingProperties().getPropertyList();
			@SuppressWarnings("deprecation")
			CTProperty[] ctPropertiesArray = cp.getUnderlyingProperties().getPropertyArray();

			for (int i = 0; i < ctPropertiesArray.length; i++) {
				if (ctPropertiesArray[i].getName().equalsIgnoreCase(key)) {
					ctProperties.remove(i);
					flag = true;
					break;
				}
			}
		}

		return flag;
	}

	@Override
	public boolean deleteTags(Collection<String> keys) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		return deleteTags(keys, false);
	}

	@Override
	public boolean deleteTags(Collection<String> keys, boolean editSecured) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException {
		CustomProperties cp = document.getProperties().getCustomProperties();
		boolean flag = false;

		if (cp != null && keys != null) {
			for (String key : keys) {
				// CTProperties need to be refresh after every loop
				List<CTProperty> ctProperties = cp.getUnderlyingProperties().getPropertyList();
				@SuppressWarnings("deprecation")
				CTProperty[] ctPropertiesArray = cp.getUnderlyingProperties().getPropertyArray();

				for (int i = 0; i < ctPropertiesArray.length; i++) {
					if (ctPropertiesArray[i].getName().equalsIgnoreCase(key)) {
						ctProperties.remove(i);
						flag = true;
						break;
					}
				}
			}
		}

		return flag;
	}

	@Override
	public void save(String filename) throws FileNotFoundException, IOException {
		OutputStream outputStream = null;

		try {
			outputStream = new FileOutputStream(new File(filename));
			document.write(outputStream);
		} finally {
			if (outputStream != null) {
				outputStream.close();
			}
		}
	}
	
	@Override
	public void save(OutputStream os) throws FileNotFoundException, IOException {
		try {
			document.write(os);
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
}
