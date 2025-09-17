namespace Test01
{
    public class UserDetails
    {
        public string? ID { get; set; }
        public string? Name { get; set; }
        public string? City { get; set; }
        public string? Country { get; set; }

        // Thêm một phương thức ToString để in ra cho đẹp
        public override string ToString()
        {
            return $"ID: {ID}, Name: {Name}, City: {City}, Country: {Country}";
        }
    }
}