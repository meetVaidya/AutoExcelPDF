class Solution {
    public boolean containsDuplicate(int[] nums) {
        int flag = nums.length;
        for (int i = 0; i < nums.length; i++) {
            for (int j = i + 1; j < nums.length; j++) {
                if (nums[j] == nums[i]) {
                    flag++;
                }
            }
        }

        if (flag != 0) {
            return true;
        } else {
            return false;
        }
    }

    public static void main(String[] args) {
        Solution sol = new Solution();
        int[] nums = {1,2,3,4};
        System.out.println(sol.containsDuplicate(nums));
    }
}