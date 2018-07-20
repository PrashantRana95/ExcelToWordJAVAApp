package com.bvicam.user.dto;


public class UserDTO {

	private int paperid;
	private double track_session;
	private String author_name;
	private String title_name;
	
	public int getPaperid() {
		return paperid;
	}
	public void setPaperid(int paperid) {
		this.paperid = paperid;
	}
	public double getTrack_session() {
		return track_session;
	}
	public void setTrack_session(double track_session) {
		this.track_session = track_session;
	}
	public String getAuthor_name() {
		return author_name;
	}
	public void setAuthor_name(String author_name) {
		this.author_name = author_name;
	}
	public String getTitle_name() {
		return title_name;
	}
	public void setTitle_name(String title_name) {
		this.title_name = title_name;
	}
	@Override
	public String toString() {
		return "UserDTO [paperid=" + paperid + ", track_session=" + track_session + ", author_name=" + author_name
				+ ", title_name=" + title_name + "]";
	}
	
}
