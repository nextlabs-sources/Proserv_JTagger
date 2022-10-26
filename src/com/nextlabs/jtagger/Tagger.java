package com.nextlabs.jtagger;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Collection;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hpsf.MarkUnsupportedException;
import org.apache.poi.hpsf.NoPropertySetStreamException;
import org.apache.poi.hpsf.UnexpectedPropertySetTypeException;
import org.apache.poi.hpsf.WritingNotSupportedException;

public interface Tagger {

	public HashMap<String, Object> getAllTags() throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException;

	public HashMap<String, Object> getAllTags(boolean bToLowerCase) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException;

	public boolean deleteAllTags() throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException;

	public boolean deleteAllTags(boolean editSecured) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException;

	public Object getTag(String key) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException;

	public boolean addTag(String key, Object value) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException;

	public boolean addTag(String key, Object value, boolean editSecured) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException;

	public boolean addTags(Map<String, Object> tags) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException;

	public boolean addTags(Map<String, Object> tags, boolean editSecured) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException;

	public boolean deleteTag(String key) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException;

	public boolean deleteTag(String key, boolean editSecured) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException;

	public boolean deleteTags(Collection<String> keys) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException;

	public boolean deleteTags(Collection<String> keys, boolean editSecured) throws IOException, NoPropertySetStreamException, MarkUnsupportedException, UnexpectedPropertySetTypeException, WritingNotSupportedException;

	public void save(String filename) throws FileNotFoundException, IOException;
	
	public void save(OutputStream os) throws FileNotFoundException, IOException;

	public void close() throws IOException;
}