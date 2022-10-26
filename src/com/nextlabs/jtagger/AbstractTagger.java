package com.nextlabs.jtagger;

public abstract class AbstractTagger implements Tagger {

	private String filename;

	public AbstractTagger(String filename) {
		super();
		this.filename = filename;
	}
	
	public AbstractTagger() {};

	public String getFilename() {
		return filename;
	}

}
