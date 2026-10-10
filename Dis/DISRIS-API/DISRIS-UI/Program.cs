using DISRIS_UI.Inferface;
using DISRIS_UI.Service;
using Microsoft.OpenApi;
using Swashbuckle.AspNetCore.Swagger;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.AddControllers();

// this is for add the swagger in api

//adding singleton 
builder.Services.AddSingleton<IAuthService, AuthService>();



builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(options =>
{
    options.SwaggerDoc("v1", new OpenApiInfo
    {
        Title = "Dis",
        Version = "v1",
        Description = "API documentation for Dis_Ris"
    });
});

var app = builder.Build();

////add swagger service for web ui
app.UseSwagger();
app.UseSwaggerUI(options =>
{
    options.SwaggerEndpoint("/swagger/v1/swagger.json", "Dis API v1");
    options.RoutePrefix = "Swagger"; // Set Swagger UI at the app's root
});



app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
