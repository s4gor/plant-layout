require 'roo'

$plant = Roo::Excelx.new("./plant.xlsx")

first_line =  [2, 1, 6, 3]
second_line = [7, 5, 8, 4]

length = first_line.length
actual_length = length * 2
total = 0

print first_line
puts ''
print second_line

puts ''

for i in 1..actual_length-1
  if first_line.include? i
    main = first_line.index(i)
  else
    main = second_line.index(i)
  end
  main += 1
  for j in i+1..actual_length
    if first_line.include? j
      second_main = first_line.index(j)
    else
      second_main = second_line.index(j)
    end
    second_main += 1

    step = (main - second_main).abs

    if step == 0
      step = 1
    end
    data = $plant.cell(i, j-1)
    puts "#{i} - #{j} -> #{data} x #{step} = #{data * step}"
    total += data * step
  end
  if i+1 != 8
    puts "new line #{i+1}"
  end
end

puts "Total: #{total}"