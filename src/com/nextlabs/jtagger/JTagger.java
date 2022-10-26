package com.nextlabs.jtagger;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class JTagger {

	public static void main(String[] args) throws Exception {
		JTagger jTagger = new JTagger();
		try {
//			String file = ("D:\\NextLabs\\Test data\\JTagger\\PDF\\Output.pdf");
//			String file = ("D:\\NextLabs\\Test data\\JTagger\\PDF\\Output.pdf");
//			String file = ("D:\\NextLabs\\Test data\\JTagger\\DOC\\Output.doc");
			String file = ("D:\\NextLabs\\Test data\\JTagger\\DOCX\\Output.docx");
//			String file = ("D:\\NextLabs\\Test data\\JTagger\\XLS\\Output.xls");
//			String file = ("D:\\NextLabs\\Test data\\JTagger\\XLSX\\Output.xlsx");
//			String file = ("D:\\NextLabs\\Test data\\JTagger\\PPT\\Output.ppt");
//			String file = ("D:\\NextLabs\\Test data\\JTagger\\PPTX\\Output.pptx");

//			jTagger.init(file);
			jTagger.addSingleTag(file);
//			jTagger.replaceSingleTag(file);
//			jTagger.removeSingleTag(file);
//			jTagger.addMultipleTags(file);
//			jTagger.replaceMultipleTags(file);
//			jTagger.removeMultipleTags(file);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void addSingleTag(String file) {
		try {
			Tagger inputTagger = TaggerFactory.getTagger(file);
			HashMap<String, Object> inputTags = inputTagger.getAllTags();

			System.out.println("================    Add Single Tag    ================");
			System.out.println("================       Original       ================");
			for (Map.Entry<String, Object> tag : inputTags.entrySet()) {
				System.out.println("Key:" + tag.getKey() + "; Value:" + tag.getValue());
			}
			System.out.println("================        Action        ================");
			System.out.println("Add Hello1 tag: " + inputTagger.addTag("Hello1", "World1"));
			System.out.println("Add Hello2 tag: " + inputTagger.addTag("Hello2", "World2"));
			System.out.println("Add Itar tag: " + inputTagger.addTag("Itar", "This is itar"));
			System.out.println("Rewrite Itar tag: " + inputTagger.addTag("Itar", "False"));

			System.out.println("Add Hello3 tag: " + inputTagger.addTag("Hello3", "World3", true));
			System.out.println("Add Hello4 tag: " + inputTagger.addTag("Hello4", "World4", true));
			System.out.println("Add Itar2 tag: " + inputTagger.addTag("Itar2", "Itar Tag 2", true));
			System.out.println("Rewrite Itar2 tag: " + inputTagger.addTag("Itar2", "True", true));
			inputTagger.save(file);

			System.out.println("================        Output        ================");
			Tagger outputTagger = TaggerFactory.getTagger(file);
			HashMap<String, Object> outputTags = outputTagger.getAllTags();

			for (Map.Entry<String, Object> tag : outputTags.entrySet()) {
				System.out.println("Key:" + tag.getKey() + "; Value:" + tag.getValue());
			}

			System.out.println("======================================================\n");
		} catch (Exception err) {
			err.printStackTrace();
		}
	}

	public void replaceSingleTag(String file) {
		try {
			Tagger inputTagger = TaggerFactory.getTagger(file);
			HashMap<String, Object> inputTags = inputTagger.getAllTags();

			System.out.println("================  Replace Single Tag  ================");
			System.out.println("================       Original       ================");
			for (Map.Entry<String, Object> tag : inputTags.entrySet()) {
				System.out.println("Key:" + tag.getKey() + "; Value:" + tag.getValue());
			}
			System.out.println("================        Action        ================");
			System.out.println("Replace Hello1 tag: " + inputTagger.addTag("Hello1", "Replaced World1"));
			System.out.println("Replace Hello2 tag: " + inputTagger.addTag("Hello2", "Replaced World2"));
			System.out.println("Replace Itar tag: " + inputTagger.addTag("Itar", "Replaced This is itar"));
			System.out.println("Rewrite Itar tag: " + inputTagger.addTag("Itar", "True"));

			System.out.println("Replace Hello3 tag: " + inputTagger.addTag("Hello3", "Replaced World3", true));
			System.out.println("Replace Hello4 tag: " + inputTagger.addTag("Hello4", "Replaced World4", true));
			System.out.println("Replace Itar2 tag: " + inputTagger.addTag("Itar2", "Replaced Itar Tag 2", true));
			System.out.println("Rewrite Itar2 tag: " + inputTagger.addTag("Itar2", "False", true));
			inputTagger.save(file);

			System.out.println("================        Output        ================");
			Tagger outputTagger = TaggerFactory.getTagger(file);
			HashMap<String, Object> outputTags = outputTagger.getAllTags();

			for (Map.Entry<String, Object> tag : outputTags.entrySet()) {
				System.out.println("Key:" + tag.getKey() + "; Value:" + tag.getValue());
			}

			System.out.println("======================================================\n");
		} catch (Exception err) {
			err.printStackTrace();
		}
	}

	public void removeSingleTag(String file) {
		try {
			Tagger inputTagger = TaggerFactory.getTagger(file);
			HashMap<String, Object> inputTags = inputTagger.getAllTags();

			System.out.println("================   Remove Single Tag  ================");
			System.out.println("================       Original       ================");
			for (Map.Entry<String, Object> tag : inputTags.entrySet()) {
				System.out.println("Key:" + tag.getKey() + "; Value:" + tag.getValue());
			}
			System.out.println("================        Action        ================");
			System.out.println("Remove Hello1 tag: " + inputTagger.deleteTag("Hello1"));
			System.out.println("Remove Hello2 tag: " + inputTagger.deleteTag("Hello2"));
			System.out.println("Remove Itar tag: " + inputTagger.deleteTag("Itar"));
			System.out.println("Remove Itar tag again: " + inputTagger.deleteTag("Itar"));
			System.out.println("Remove NotExist1 tag: " + inputTagger.deleteTag("NotExist1"));

			System.out.println("Remove Hello3 tag: " + inputTagger.deleteTag("Hello3", true));
			System.out.println("Remove Hello4 tag: " + inputTagger.deleteTag("Hello4", true));
			System.out.println("Remove Itar2 tag: " + inputTagger.deleteTag("Itar2", true));
			System.out.println("Remove Itar2 tag again: " + inputTagger.deleteTag("Itar2", true));
			System.out.println("Remove NotExist2 tag: " + inputTagger.deleteTag("NotExist2", true));
			inputTagger.save(file);

			System.out.println("================        Output        ================");
			Tagger outputTagger = TaggerFactory.getTagger(file);
			HashMap<String, Object> outputTags = outputTagger.getAllTags();

			for (Map.Entry<String, Object> tag : outputTags.entrySet()) {
				System.out.println("Key:" + tag.getKey() + "; Value:" + tag.getValue());
			}

			System.out.println("======================================================\n");
		} catch (Exception err) {
			err.printStackTrace();
		}
	}

	public void addMultipleTags(String file) {
		try {
			Tagger inputTagger = TaggerFactory.getTagger(file);
			HashMap<String, Object> inputTags = inputTagger.getAllTags();

			System.out.println("================   Add Multiple Tags  ================");
			System.out.println("================       Original       ================");
			for (Map.Entry<String, Object> tag : inputTags.entrySet()) {
				System.out.println("Key:" + tag.getKey() + "; Value:" + tag.getValue());
			}
			System.out.println("================        Action        ================");
			HashMap<String, Object> tags1 = new HashMap<String, Object>();
			tags1.put("Tag1", "Value 1");
			tags1.put("Tag2", "Value 2");
			tags1.put("Tag3", "Value 3");

			System.out.println("Add multiple tags: " + inputTagger.addTags(tags1));
			HashMap<String, Object> tags2 = new HashMap<String, Object>();
			tags2.put("Tag4", "Value 4");
			tags2.put("Tag5", "Value 5");
			tags2.put("Tag6", "Value 6");
			System.out.println("Add multiple tags: " + inputTagger.addTags(tags2, true));
			inputTagger.save(file);

			System.out.println("================        Output        ================");
			Tagger outputTagger = TaggerFactory.getTagger(file);
			HashMap<String, Object> outputTags = outputTagger.getAllTags();

			for (Map.Entry<String, Object> tag : outputTags.entrySet()) {
				System.out.println("Key:" + tag.getKey() + "; Value:" + tag.getValue());
			}

			System.out.println("======================================================\n");
		} catch (Exception err) {
			err.printStackTrace();
		}
	}

	public void replaceMultipleTags(String file) {
		try {
			Tagger inputTagger = TaggerFactory.getTagger(file);
			HashMap<String, Object> inputTags = inputTagger.getAllTags();

			System.out.println("================ Replace Multiple Tag ================");
			System.out.println("================       Original       ================");
			for (Map.Entry<String, Object> tag : inputTags.entrySet()) {
				System.out.println("Key:" + tag.getKey() + "; Value:" + tag.getValue());
			}
			System.out.println("================        Action        ================");
			HashMap<String, Object> tags1 = new HashMap<String, Object>();
			tags1.put("Tag1", "Updated Value 1");
			tags1.put("Tag2", "Updated Value 2");
			tags1.put("Tag3", "Updated Value 3");

			System.out.println("Add multiple tags: " + inputTagger.addTags(tags1));
			HashMap<String, Object> tags2 = new HashMap<String, Object>();
			tags2.put("Tag4", "Updated Value 4");
			tags2.put("Tag5", "Updated Value 5");
			tags2.put("Tag6", "Updated Value 6");
			System.out.println("Add multiple tags: " + inputTagger.addTags(tags2, true));
			inputTagger.save(file);

			System.out.println("================        Output        ================");
			Tagger outputTagger = TaggerFactory.getTagger(file);
			HashMap<String, Object> outputTags = outputTagger.getAllTags();

			for (Map.Entry<String, Object> tag : outputTags.entrySet()) {
				System.out.println("Key:" + tag.getKey() + "; Value:" + tag.getValue());
			}

			System.out.println("======================================================\n");
		} catch (Exception err) {
			err.printStackTrace();
		}
	}

	public void removeMultipleTags(String file) {
		try {
			Tagger inputTagger = TaggerFactory.getTagger(file);
			HashMap<String, Object> inputTags = inputTagger.getAllTags();

			System.out.println("================ Remove Multiple Tag  ================");
			System.out.println("================       Original       ================");
			for (Map.Entry<String, Object> tag : inputTags.entrySet()) {
				System.out.println("Key:" + tag.getKey() + "; Value:" + tag.getValue());
			}
			System.out.println("================        Action        ================");
			List<String> tags1 = new ArrayList<String>();
			tags1.add("Tag1");
			tags1.add("Tag2");
			tags1.add("Tag3");
			tags1.add("NotExist1");

			System.out.println("Remove multiple tags: " + inputTagger.deleteTags(tags1));
			Set<String> tags2 = new HashSet<String>();
			tags2.add("Tag4");
			tags2.add("Tag5");
			tags2.add("Tag6");
			tags2.add("NotExist2");
			System.out.println("Remove multiple tags: " + inputTagger.deleteTags(tags2, true));
			inputTagger.save(file);

			System.out.println("================        Output        ================");
			Tagger outputTagger = TaggerFactory.getTagger(file);
			HashMap<String, Object> outputTags = outputTagger.getAllTags();

			for (Map.Entry<String, Object> tag : outputTags.entrySet()) {
				System.out.println("Key:" + tag.getKey() + "; Value:" + tag.getValue());
			}

			System.out.println("======================================================\n");
		} catch (Exception err) {
			err.printStackTrace();
		}
	}

	public void removeAllTags(String file) {
		try {
			Tagger inputTagger = TaggerFactory.getTagger(file);
			HashMap<String, Object> inputTags = inputTagger.getAllTags();

			System.out.println("================   Remove All Tags    ================");
			System.out.println("================       Original       ================");
			for (Map.Entry<String, Object> tag : inputTags.entrySet()) {
				System.out.println("Key:" + tag.getKey() + "; Value:" + tag.getValue());
			}
			System.out.println("================        Action        ================");
			System.out.println("Remove all tags: " + inputTagger.deleteAllTags());
			inputTagger.save(file);

			System.out.println("================        Output        ================");
			Tagger outputTagger = TaggerFactory.getTagger(file);
			HashMap<String, Object> outputTags = outputTagger.getAllTags();

			for (Map.Entry<String, Object> tag : outputTags.entrySet()) {
				System.out.println("Key:" + tag.getKey() + "; Value:" + tag.getValue());
			}

			System.out.println("======================================================\n");
		} catch (Exception err) {
			err.printStackTrace();
		}
	}

	public void init(String file) {
		try {
			Tagger inputTagger = TaggerFactory.getTagger(file);
			inputTagger.deleteAllTags(true);
			inputTagger.addTag("DocVer", "v2.0.4", true);
			inputTagger.addTag("DocTag", "T1.0", true);

			inputTagger.save(file);
		} catch (Exception err) {
			err.printStackTrace();
		}
	}
}
