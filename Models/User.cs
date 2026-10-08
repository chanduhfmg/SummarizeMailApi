using System;
using System.Collections.Generic;

namespace HorizonChat.Models;

public partial class User
{
    public int Id { get; set; }

    public string Account { get; set; } = null!;

    public string Name { get; set; } = null!;

    public string Mail { get; set; } = null!;

    public string? Role { get; set; }

    public string? Password { get; set; }

    public DateTime? Created { get; set; }
}
