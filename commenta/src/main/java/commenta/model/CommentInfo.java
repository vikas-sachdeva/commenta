package commenta.model;

public class CommentInfo {

	private String comment;

	private int rowNo;

	private int columnNo;

	private String commentedText;

	private String commentAuthor;

	public String getCommentAuthor() {
		return commentAuthor;
	}

	public void setCommentAuthor(String commentAuthor) {
		this.commentAuthor = commentAuthor;
	}

	public String getComment() {
		return comment;
	}

	public void setComment(String comment) {
		this.comment = comment;
	}

	public int getRowNo() {
		return rowNo;
	}

	public void setRowNo(int rowNo) {
		this.rowNo = rowNo;
	}

	public int getColumnNo() {
		return columnNo;
	}

	public void setColumnNo(int columnNo) {
		this.columnNo = columnNo;
	}

	public String getCommentedText() {
		return commentedText;
	}

	public void setCommentedText(String commentedText) {
		this.commentedText = commentedText;
	}
}
