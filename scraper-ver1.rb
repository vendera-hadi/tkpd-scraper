# frozen_string_literal: true

require 'nokogiri'
require 'spreadsheet'
require 'watir'
require 'byebug'

def current_spreadsheet
  @current_spreadsheet ||= Spreadsheet.open('tokped.xls')
end

def category_sheet
  @category_sheet ||= begin
    result = {} 
    list = current_spreadsheet.worksheet(1)
    list.each do |row|
      result[row[2]] = row[3].to_i
    end
    result
  end
end

def content_sheet
  current_spreadsheet.worksheet(0)
end

def new_row_index
  content_sheet.last_row_index + 1
end

def imgur_client
  @imgur_client ||= Imgur.new('2c8bcbf4afd0d38')
end

def scrap_page(url)
  browser = Watir::Browser.new
  browser.goto url

  # scroll down & wait 2 seconds
  browser.scroll.to :bottom
  sleep(2)

  Nokogiri::HTML.parse(browser.html)
end

def image_big(url)
  url.gsub('100-square', '500-square')
end

def generate_desc(original_desc)
  additional = "<br><br>Semua barang ready stock yah.<br> 
  Pemesanan hari senin - sabtu sebelum pukul 16.00 akan kami kirim di hari yang sama dan diatas itu dikirim keesokan harinya.<br> 
  Pemesanan hari Minggu dan tanggal merah akan diproses pada hari kerja berikutnya.<br> 
  <br>
  Kami menjual produk spare part maupun accecories serta kebutuhan untuk offroad khusus Jeep.<br> 
  Silahkan chat kami bila ada pertanyaaan mengenai produk yang sudah ada di list produk maupun yang belum ada di list produk kami.<br> 
  <br>
  Selamat berbelanja di toko kami 🙏"
  "#{original_desc}#{additional}".dup.gsub!("<br>", "\x0A")
end

def price_with_margin(price)
  percentage = 50 if price <= 50_000
  percentage = 30 if price > 50_000 && price <= 100_000
  percentage = 20 if price > 100_000 && price <= 1_000_000
  percentage = 10 if price > 1_000_000 && price <= 2_500_000
  percentage = 7.5 if price > 2_500_000 && price <= 5_000_000
  percentage = 5 if price > 5_000_000
  percentage_val = percentage.to_f / 100
  price.to_f + (percentage_val * price.to_f)
end

def fetch_page(page)
  # variables
  name = page.css('.css-v7vvdw').text
  desc = generate_desc(page.css('.css-168ydy0 div').text)
  desc_detail = page.css('.css-r0v3c0')
  cat_text = desc_detail[2].css('a').text
  cat_id = category_sheet[cat_text]
  weight_text = desc_detail[1].css('span.main').text
  weight = weight_text.tr('^0-9', '').to_i
  condition = desc_detail[0].css('span.main').text
  price_text = page.css('.css-32gaxy .price').text
  price = price_text.tr('^0-9', '').to_i
  selling_price = price_with_margin(price)
  stock = page.css('.css-9vy0ue p b').text
  active = stock.to_i.zero? ? 'Non Aktif' : 'Aktif'
  etalase_no = 27355886
  images = page.css('.css-ikbtkl').map{|container| image_big(container.css('img').first['src']) }
  min_order = 1
  asuransi = 'opsional'

  # print results
  puts "Nama : #{name}"
  puts "Deskripsi : #{desc}"
  puts "Kategori Kode : #{cat_id}"
  puts "Berat (gr) : #{weight}"
  puts "Minimum Order : #{min_order}"
  puts "Nomor Etalase : #{etalase_no}"
  puts "Kondisi : #{condition}"
  puts "Gambar (X): #{images}"
  puts "Status : #{active}"
  puts "Jumlah Stok : #{stock}"
  puts "Harga : #{selling_price}"
  puts "Asuransi Pengiriman : #{asuransi}"

  # save results to excel
  new_row = [nil, name, desc, cat_id, weight, min_order, etalase_no, nil, condition, images[0], images[1], images[2], images[3], images[4], nil, nil, nil, nil, active, stock, selling_price, asuransi]
  content_sheet.insert_row(new_row_index, new_row)
end

def perform
  total_scrapped = 0
  total_products = 1128
  per_page = 10
  # total_scraps = (total_products.to_f/per_page.to_f).ceil
  total_scraps = 2

  # loop every page
  total_scraps.times do |index|
    page = index + 1
    puts '=' * 10
    puts "Page #{page}"
    puts '=' * 10
    url = "https://www.tokopedia.com/parts-jeep/page/#{page}?perpage=#{per_page}"
    pagination_parsed_page = scrap_page(url)
    products = pagination_parsed_page.css('.css-1sn1xa2')
    # loop every product in one page
    products.each_with_index do |product, product_idx|
      next if product_idx < 4
      product_url = product.css('.css-1ehqh5q a').first['href']
      puts "Product URL: #{product_url}"
      parsed_page = scrap_page(product_url)
      fetch_page(parsed_page)
      total_scrapped += 1
      puts '===================='
      break
    end
    break
  end

  current_spreadsheet.write("tokped_#{Time.now.strftime('%y%m%d')}.xls")
  puts "Total Scrapped : #{total_scrapped} products"
end

perform
