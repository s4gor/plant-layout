require 'roo'

$plant = Roo::Excelx.new("./plant.xlsx")


parent = []


for i in 1..7
  parent[i] = []
  for j in 1..7
    parent[i][j] = $plant.cell(i, j)
  end
end

print parent