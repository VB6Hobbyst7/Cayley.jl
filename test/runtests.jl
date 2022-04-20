using Cayley
using Test

@testset "Cayley.jl" begin
    @test Cayley.interpcreditline([1, 2, 3, 4, 5, 7, 10], [350, 350, 350, 250, 200, 70, 15],
        [0, 3, 3.1, 11], "FlatToRight") == [350, 350, 250, 15]
    @test Cayley.interpcreditline([1, 2, 3, 4, 5, 7, 10], [350, 350, 350, 250, 200, 70, 15],
        [0, 1.5, 3.1, 11], "Linear") ≈ [350, 350, 340, 15]
end
