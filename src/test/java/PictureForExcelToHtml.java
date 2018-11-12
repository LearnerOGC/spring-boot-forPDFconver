public class PictureForExcelToHtml {
    private int left_up_x;
    private int left_up_y;
    private int right_down_x;
    private int right_down_y;

    public PictureForExcelToHtml(int left_up_x, int left_up_y, int right_down_x, int right_down_y) {
        this.left_up_x = left_up_x;
        this.left_up_y = left_up_y;
        this.right_down_x = right_down_x;
        this.right_down_y = right_down_y;
    }

    public int getLeft_up_x() {
        return left_up_x;
    }

    public void setLeft_up_x(int left_up_x) {
        this.left_up_x = left_up_x;
    }

    public int getLeft_up_y() {
        return left_up_y;
    }

    public void setLeft_up_y(int left_up_y) {
        this.left_up_y = left_up_y;
    }

    public int getRight_down_x() {
        return right_down_x;
    }

    public void setRight_down_x(int right_down_x) {
        this.right_down_x = right_down_x;
    }

    public int getRight_down_y() {
        return right_down_y;
    }

    public void setRight_down_y(int right_down_y) {
        this.right_down_y = right_down_y;
    }
}
